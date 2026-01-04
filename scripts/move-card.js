define(["require", "exports", "TFS/Core/RestClient", "TFS/Work/RestClient", "TFS/WorkItemTracking/RestClient"], 
function (require, exports, CoreRestClient, WorkRestClient, WitRestClient) {
    "use strict";

    // Registrar o objeto corretamente para evitar o erro "undefined"
    var provider = {
        // Se precisar de callbacks do form no futuro
    };
    
    // Registra usando o ID dinâmico da contribuição para evitar falhas
    VSS.register(VSS.getContribution().id, provider);

    // --- CONFIGURAÇÃO ---
    var $boardSelect = document.getElementById("boardSelect");
    var $columnSelect = document.getElementById("columnSelect");
    var $moveBtn = document.getElementById("moveBtn");
    var $status = document.getElementById("status");
    
    var context = VSS.getWebContext();
    var currentColumns = [];

    function logStatus(msg, type) {
        console.log("[move-card] " + msg);
        if ($status) {
            $status.innerText = msg; // Usar innerText é mais seguro que textContent as vezes
            $status.className = type || "";
        }
    }

    // Função auxiliar para tratamento de erros
    function handleError(err, contextMsg) {
        console.error("[move-card-error] " + contextMsg, err);
        var msg = err.message || (err.responseText ? err.responseText : "Erro desconhecido");
        
        if (msg.indexOf("401") !== -1 || msg.indexOf("Unauthorized") !== -1) {
            logStatus("Erro de Permissão (401). Verifique se o usuário tem acesso aos Boards.", "error");
        } else if (msg.indexOf("Access-Control-Allow-Origin") !== -1) {
            logStatus("Erro de CORS (Rede). O servidor bloqueou a requisição.", "error");
        } else {
            logStatus("Erro: " + msg, "error");
        }
    }

    // 1. Carregar Times
    function loadTeams() {
        logStatus("Autenticando e carregando times...", "loading");
        
        // Obtém o cliente HTTP. O VSS injeta o token automaticamente aqui.
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
            
            logStatus("Times carregados.", "success");

            if (context.team && context.team.id) {
                $boardSelect.value = context.team.id;
                loadColumns(context.team.id);
            }
        }, function(err) {
            handleError(err, "Falha ao carregar times");
        });
    }

    // 2. Carregar Colunas (Onde estava dando erro 401)
    function loadColumns(teamId) {
        if (!teamId) return;

        $columnSelect.innerHTML = '<option>Carregando colunas...</option>';
        $columnSelect.disabled = true;
        $moveBtn.disabled = true;

        var workClient = WorkRestClient.getClient();

        // Passamos explicitamente o project e team ID
        var teamContext = { projectId: context.project.id, teamId: teamId, project: context.project.name, team: "" };

        console.log("[move-card] Buscando boards para o time ID: " + teamId);

        workClient.getBoards(teamContext).then(function(boards) {
            if (!boards || boards.length === 0) {
                $columnSelect.innerHTML = '<option>Nenhum board encontrado</option>';
                logStatus("Nenhum board encontrado neste time.", "error");
                return;
            }

            // Pega o primeiro board disponível
            var targetBoard = boards[0];
            console.log("[move-card] Board encontrado: " + targetBoard.name + " (" + targetBoard.id + ")");

            workClient.getBoard(teamContext, targetBoard.id).then(function(boardDetails) {
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
                
            }, function(err) {
                handleError(err, "Falha ao pegar detalhes do board");
            });

        }, function(err) {
            handleError(err, "Falha ao listar boards (401 aqui significa falta de escopo vso.work)");
        });
    }

    // 3. Mover Card
    function moveCard() {
        $moveBtn.disabled = true;
        logStatus("Aguarde...", "loading");

        VSS.getService(VSS.ServiceIds.WorkItemFormService).then(function(workItemFormService) {
            workItemFormService.getId().then(function(id) {
                if (!id) {
                    logStatus("Salve o Work Item primeiro.", "error");
                    $moveBtn.disabled = false;
                    return;
                }

                // Pega o tipo para mapear o estado
                workItemFormService.getFieldValue("System.WorkItemType").then(function(wiType) {
                    
                    var colId = $columnSelect.value;
                    var targetCol = currentColumns.find(function(c) { return c.id === colId; });
                    
                    if (!targetCol) {
                        logStatus("Coluna inválida.", "error");
                        $moveBtn.disabled = false;
                        return;
                    }

                    // Lógica de Estado
                    var targetState = targetCol.name; 
                    if (targetCol.stateMappings && targetCol.stateMappings[wiType]) {
                        targetState = targetCol.stateMappings[wiType];
                    }

                    console.log("[move-card] Movendo ID " + id + " para Estado: " + targetState);

                    var patchDocument = [
                        { "op": "add", "path": "/fields/System.State", "value": targetState },
                        { "op": "add", "path": "/fields/System.History", "value": "Movido via Extensão" }
                    ];

                    var witClient = WitRestClient.getClient();
                    
                    witClient.updateWorkItem(patchDocument, id).then(function(updated) {
                        logStatus("Sucesso! Recarregue a página.", "success");
                        $moveBtn.disabled = false;
                        
                        // Forçar refresh no Azure DevOps On-Prem
                        VSS.getService(VSS.ServiceIds.NavigationService).then(function(navigationService) {
                             if(navigationService.reload) {
                                 navigationService.reload();
                             } else {
                                 window.location.reload();
                             }
                        });

                    }, function(err) {
                        handleError(err, "Falha ao atualizar Work Item");
                        $moveBtn.disabled = false;
                    });
                });
            });
        }, function(err) {
            handleError(err, "Falha ao conectar no FormService");
        });
    }

    // Eventos
    $boardSelect.addEventListener("change", function() { loadColumns(this.value); });
    $moveBtn.addEventListener("click", moveCard);

    // Inicialização
    loadTeams();
    
    // Notifica que terminou de carregar
    VSS.notifyLoadSucceeded();
});