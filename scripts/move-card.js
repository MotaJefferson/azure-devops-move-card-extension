define(["require", "exports", "TFS/Core/RestClient", "TFS/Work/RestClient", "TFS/WorkItemTracking/RestClient"], 
function (require, exports, CoreRestClient, WorkRestClient, WitRestClient) {
    "use strict";

    // --- VARIÁVEIS GLOBAIS ---
    var provider = {};
    // Registra a extensão
    VSS.register(VSS.getContribution().id, provider);

    var $boardSelect = document.getElementById("boardSelect");
    var $columnSelect = document.getElementById("columnSelect");
    var $moveBtn = document.getElementById("moveBtn");
    var $status = document.getElementById("status");
    
    var context = VSS.getWebContext();
    var currentColumns = []; 
    var targetTeamSettings = {
        areaPath: null,
        columnFieldRef: null
    };

    // --- FUNÇÕES UTILITÁRIAS ---
    function logStatus(msg, type) {
        console.log("[move-card] " + msg);
        if ($status) {
            $status.innerText = msg;
            $status.className = type || "";
        }
    }

    function handleError(err, contextMsg) {
        console.error("[move-card-error] " + contextMsg, err);
        var msg = err.message || (err.responseText ? err.responseText : "Erro desconhecido");
        
        // Tratamento simples de erros comuns
        if (msg.indexOf("HostAuthorizationNotFound") !== -1) {
            msg = "Erro de Permissão (500). Tente reinstalar a extensão.";
        }

        logStatus(msg, "error");
        $moveBtn.disabled = false;
    }

    // --- LÓGICA ---

    function loadTeams() {
        logStatus("Carregando times...", "loading");
        var coreClient = CoreRestClient.getClient();
        
        coreClient.getTeams(context.project.id).then(function(teams) {
            $boardSelect.innerHTML = '<option value="">Selecione um time</option>';
            teams.sort(function(a, b) { return a.name.localeCompare(b.name); });
            teams.forEach(function(t) {
                var opt = document.createElement("option");
                opt.value = t.id;
                opt.text = t.name;
                $boardSelect.appendChild(opt);
            });
            logStatus("Aguardando seleção...", "");
        }, function(err) {
            handleError(err, "Falha ao carregar times");
        });
    }

    function loadBoardDetails(teamId) {
        if (!teamId) {
            $columnSelect.innerHTML = '<option>—</option>';
            $columnSelect.disabled = true;
            $moveBtn.disabled = true;
            return;
        }

        $columnSelect.innerHTML = '<option>Carregando...</option>';
        $columnSelect.disabled = true;
        $moveBtn.disabled = true;
        
        var workClient = WorkRestClient.getClient();
        var teamContext = { projectId: context.project.id, teamId: teamId };

        Promise.all([
            workClient.getBoards(teamContext),
            workClient.getTeamFieldValues(teamContext)
        ]).then(function(results) {
            var boards = results[0];
            var fieldValues = results[1];

            targetTeamSettings.areaPath = fieldValues.defaultValue;

            if (!boards || boards.length === 0) throw new Error("Sem boards neste time.");
            var targetBoard = boards[0]; 

            return workClient.getBoard(teamContext, targetBoard.id);

        }).then(function(boardDetails) {
            
            // Detecta o campo de coluna (WEF)
            if (boardDetails.fields && boardDetails.fields.columnField) {
                targetTeamSettings.columnFieldRef = boardDetails.fields.columnField.referenceName;
            } else {
                targetTeamSettings.columnFieldRef = null; 
            }

            currentColumns = boardDetails.columns;
            $columnSelect.innerHTML = ''; 
            
            if (!currentColumns || currentColumns.length === 0) {
                 $columnSelect.innerHTML = '<option>Sem colunas</option>';
                 return;
            }

            currentColumns.forEach(function(col) {
                var opt = document.createElement("option");
                opt.value = col.id;
                opt.text = col.name;
                $columnSelect.appendChild(opt);
            });

            $columnSelect.disabled = false;
            $moveBtn.disabled = false;
            logStatus("Pronto.", "success");

        }).catch(function(err) {
            handleError(err, "Falha ao carregar board");
        });
    }

    // 3. Executar Movimento
    function executeMove() {
        $moveBtn.disabled = true;
        logStatus("Movendo...", "loading");

        VSS.getService("ms.vss-work-web.work-item-form").then(function(workItemFormService) {
            
            workItemFormService.getId().then(function(id) {
                if (!id) { logStatus("Salve o item antes.", "error"); return; }

                workItemFormService.getFieldValues(["System.WorkItemType", "System.AreaPath", "System.BoardColumn"])
                .then(function(currentValues) {

                    var type = currentValues["System.WorkItemType"];
                    var sourceArea = currentValues["System.AreaPath"];
                    var sourceColumn = currentValues["System.BoardColumn"] || "Sem Coluna";

                    // === LÓGICA DA MENSAGEM (RESTAURADA) ===
                    // 1. Extrai apenas o nome do time da Área (ex: "Projeto\Time 01" vira "Time 01")
                    var sourceTeamName = sourceArea.split('\\').pop();

                    // 2. Pega o nome do time destino
                    var targetTeamName = $boardSelect.options[$boardSelect.selectedIndex].text;
                    
                    // 3. Pega a coluna destino
                    var colId = $columnSelect.value;
                    var targetCol = currentColumns.find(function(c) { return c.id === colId; });
                    var targetColumnName = targetCol ? targetCol.name : "Desconhecida";

                    // Mapeia Estado
                    var targetState = targetCol.name; 
                    if (targetCol && targetCol.stateMappings && targetCol.stateMappings[type]) {
                        targetState = targetCol.stateMappings[type];
                    }

                    // 4. Monta a mensagem formatada
                    var historyMessage = "Movido do " + sourceTeamName + " | " + sourceColumn + 
                                         " para " + targetTeamName + " | " + targetColumnName;

                    // Monta o Patch JSON
                    var patchDocument = [
                        { "op": "add", "path": "/fields/System.AreaPath", "value": targetTeamSettings.areaPath },
                        { "op": "add", "path": "/fields/System.State", "value": targetState },
                        { "op": "add", "path": "/fields/System.History", "value": historyMessage }
                    ];

                    // Adiciona coluna customizada se existir o campo WEF
                    if (targetTeamSettings.columnFieldRef) {
                        patchDocument.push({
                            "op": "add", 
                            "path": "/fields/" + targetTeamSettings.columnFieldRef, 
                            "value": targetColumnName 
                        });
                    }

                    var witClient = WitRestClient.getClient();
                    
                    witClient.updateWorkItem(patchDocument, id).then(function(updated) {
                        logStatus("Sucesso!", "success");
                        
                        // Atualiza o formulário sem recarregar a página inteira
                        workItemFormService.refresh().then(function() {
                            console.log("[move-card] Atualizado.");
                        });
                        
                        setTimeout(function(){ $moveBtn.disabled = false; }, 1500);

                    }, function(err) {
                        handleError(err, "Erro ao atualizar");
                    });
                });
            });
        });
    }

    $boardSelect.addEventListener("change", function() { loadBoardDetails(this.value); });
    $moveBtn.addEventListener("click", executeMove);

    loadTeams();
    VSS.notifyLoadSucceeded();
});