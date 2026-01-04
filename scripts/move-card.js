/* globals VSS */
/*
  move-card.js - versão final com fallback REST robusto
  - Registra contribution id: JeffersonMota.move-card-onprem.move-card-form-group
  - Usa SDK RestClients quando disponíveis
  - Fallback: tenta múltiplos endpoints REST com fetch(credentials:'same-origin')
  - Atualiza System.State via WorkItemTracking RestClient (ou PATCH fetch se necessario)
*/

(function () {
  'use strict';

  VSS.init({ explicitNotifyLoaded: true, usePlatformScripts: true });

  // Utilitários
  function log(...args) { console.log('[move-card]', ...args); }
  function warn(...args) { console.warn('[move-card]', ...args); }
  function error(...args) { console.error('[move-card]', ...args); }

  function setStatus($statusEl, text, isError) {
    if ($statusEl) {
      $statusEl.textContent = text;
      $statusEl.style.color = isError ? '#b00020' : '#333';
    }
    log('status:', text);
  }

  // Fetch com credentials (mesma origem)
  async function fetchWithCredentials(url, opts = {}) {
    const finalOpts = Object.assign({}, opts, { credentials: 'same-origin' });
    log('fetch', url, finalOpts);
    const resp = await fetch(url, finalOpts);
    return resp;
  }

  // Constrói base da collection (ex: http://jefferson/DefaultCollection/)
  function collectionBase(webContext) {
    return webContext.collection.uri.replace(/\/+$/, '') + '/';
  }

  // --- Controller principal (cria UI e lógica) ---
  function createController(root) {
    const $board = root.querySelector('#boardSelect');
    const $column = root.querySelector('#columnSelect');
    const $moveBtn = root.querySelector('#moveBtn');
    const $status = root.querySelector('#status');

    const webContext = VSS.getWebContext();
    let workItemId = null;
    let boardsCache = [];
    let columnsCache = [];

    // Obtém workItemId via WorkItemFormService
    async function ensureWorkItemId() {
      try {
        const wiService = await VSS.getService(VSS.ServiceIds.WorkItemFormService);
        const id = await wiService.getId();
        workItemId = id;
        log('workItem id =', id);
        return id;
      } catch (err) {
        throw new Error('Erro ao obter id do work item: ' + err);
      }
    }

    // --- LISTAR TEAMS (Boards) ---
    // Tenta usar Core RestClient; se falhar, tenta endpoints REST
    function loadBoards() {
      $board.disabled = true;
      $board.innerHTML = '<option>Carregando...</option>';
      setStatus($status, 'Listando teams (boards)...', false);

      // Tenta Core RestClient
      VSS.require(['TFS/Core/RestClient'], function (CoreRestClient) {
        try {
          const coreClient = CoreRestClient.getClient();
          const projectId = webContext.project.id;
          coreClient.getTeams(projectId).then(function (teams) {
            boardsCache = (teams || []).map(t => ({ id: t.id, name: t.name || t.displayName || t.id }));
            populateBoardSelect();
            setStatus($status, 'Boards carregados (' + boardsCache.length + ')', false);
            // carrega colunas do primeiro
            if (boardsCache.length) loadColumnsForTeam(boardsCache[0].id, boardsCache[0].name);
          }).catch(function (err) {
            warn('CoreRestClient.getTeams failed, fallback to REST:', err);
            fallbackLoadBoards();
          });
        } catch (e) {
          warn('CoreRestClient error, fallback:', e);
          fallbackLoadBoards();
        }
      }, function (err) {
        warn('CoreRestClient not available, fallback:', err);
        fallbackLoadBoards();
      });
    }

    function populateBoardSelect() {
      $board.innerHTML = '';
      if (!boardsCache || !boardsCache.length) {
        $board.innerHTML = '<option>(nenhum board)</option>';
        $board.disabled = true;
        return;
      }
      boardsCache.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b.id;
        opt.textContent = b.name;
        $board.appendChild(opt);
      });
      $board.disabled = false;
    }

    // Fallback teams via REST. Tenta algumas URLs comuns.
    async function fallbackLoadBoards() {
      setStatus($status, 'Fallback: buscando teams via REST...', false);
      const base = collectionBase(webContext);
      const projectName = encodeURIComponent(webContext.project.name);

      const urls = [
        // tentar endpoints típicos (vários formatos usados por versões On-Prem)
        base + '_apis/projects/' + projectName + '/teams?api-version=5.1',
        base + projectName + '/_apis/teams?api-version=5.1',
        base + '_apis/teams?api-version=5.1'
      ];

      let lastErr = null;
      for (const url of urls) {
        try {
          const res = await fetchWithCredentials(url);
          if (!res.ok) {
            const txt = await res.text().catch(()=>'');
            lastErr = `HTTP ${res.status} ${res.statusText} - ${txt}`;
            warn('fallbackLoadBoards non-ok', url, lastErr);
            continue;
          }
          const json = await res.json();
          const teams = json.value || [];
          if (!teams.length) {
            lastErr = 'Nenhuma team retornada de ' + url;
            continue;
          }
          boardsCache = teams.map(t => ({ id: t.id, name: t.name || t.displayName || t.id }));
          populateBoardSelect();
          setStatus($status, 'Boards carregados (fallback) via REST', false);
          // carrega colunas do primeiro
          await loadColumnsForTeam(boardsCache[0].id, boardsCache[0].name);
          return;
        } catch (err) {
          warn('fallbackLoadBoards error for', url, err);
          lastErr = err;
        }
      }
      $board.innerHTML = '<option>Erro</option>';
      setStatus($status, 'Erro ao listar teams (fallback). Último erro: ' + lastErr, true);
    }

    // --- LISTAR COLUNAS PARA UM TEAM ---
    function loadColumnsForTeam(teamId, teamName) {
      $column.disabled = true;
      $column.innerHTML = '<option>Carregando...</option>';
      setStatus($status, 'Listando colunas para team: ' + (teamName || teamId), false);
      columnsCache = [];

      // Tenta Work RestClient primeiro
      VSS.require(['TFS/Work/RestClient'], function (WorkRestClient) {
        try {
          const workClient = WorkRestClient.getClient();
          // API moderna: getBoardsForProject(projectId, teamId)
          if (workClient.getBoardsForProject) {
            workClient.getBoardsForProject(webContext.project.id, teamId).then(function (boards) {
              if (!boards || !boards.length) {
                warn('getBoardsForProject retornou vazio, fallback to REST');
                fallbackLoadColumns(teamId, teamName);
                return;
              }
              const board = boards[0];
              const cols = board.columns || board.columnOptions || [];
              if (!cols || !cols.length) {
                warn('board sem columns; fallback');
                fallbackLoadColumns(teamId, teamName);
                return;
              }
              columnsCache = cols;
              populateColumnSelect();
              setStatus($status, 'Colunas carregadas (' + columnsCache.length + ')', false);
            }).catch(function (err) {
              warn('getBoardsForProject erro, fallback', err);
              fallbackLoadColumns(teamId, teamName);
            });
          } else {
            warn('WorkRestClient não expõe getBoardsForProject, fallback');
            fallbackLoadColumns(teamId, teamName);
          }
        } catch (e) {
          warn('WorkRestClient error, fallback', e);
          fallbackLoadColumns(teamId, teamName);
        }
      }, function (err) {
        warn('WorkRestClient não disponível, fallback', err);
        fallbackLoadColumns(teamId, teamName);
      });
    }

    function populateColumnSelect() {
      $column.innerHTML = '';
      if (!columnsCache || !columnsCache.length) {
        $column.innerHTML = '<option>—</option>';
        $column.disabled = true;
        return;
      }
      columnsCache.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c.id || c.name || JSON.stringify(c);
        opt.textContent = c.name || c.displayName || c.id || 'Coluna';
        $column.appendChild(opt);
      });
      $column.disabled = false;
    }

    // Fallback para buscar boards/columns via REST com vários endpoints
    async function fallbackLoadColumns(teamId, teamName) {
      setStatus($status, 'Fallback: buscando boards/colunas via REST...', false);
      const base = collectionBase(webContext);
      const projectName = encodeURIComponent(webContext.project.name);
      const encTeam = encodeURIComponent(teamId || teamName || '');

      // URLs a tentar (vários formatos)
      const urls = [
        // team-specific
        `${base}${projectName}/${encTeam}/_apis/work/boards?api-version=5.1`,
        `${base}${projectName}/${encTeam}/_apis/work/boards?api-version=6.0`,
        // project-level
        `${base}${projectName}/_apis/work/boards?api-version=5.1`,
        `${base}${projectName}/_apis/work/boards?api-version=6.0`,
        // collection-level (menos provável)
        `${base}_apis/work/boards?api-version=5.1`,
        `${base}_apis/work/boards?api-version=6.0`
      ];

      let lastErr = null;
      for (const url of urls) {
        try {
          const res = await fetchWithCredentials(url);
          if (!res.ok) {
            const txt = await res.text().catch(()=>'');
            lastErr = `HTTP ${res.status} ${res.statusText} - ${txt}`;
            warn('fallbackLoadColumns non-ok', url, lastErr);
            continue;
          }
          const json = await res.json();
          const boards = json.value || [];
          if (!boards.length) {
            lastErr = 'Nenhum board no resultado de ' + url;
            continue;
          }
          const board = boards[0];
          const cols = board.columns || board.columnOptions || board.columnsResponse || [];
          if (!cols || !cols.length) {
            lastErr = 'Nenhuma coluna encontrada no board retornado por ' + url;
            continue;
          }
          columnsCache = cols;
          populateColumnSelect();
          setStatus($status, 'Colunas carregadas (fallback) via REST', false);
          return;
        } catch (err) {
          warn('fallbackLoadColumns error for', url, err);
          lastErr = err;
        }
      }

      $column.innerHTML = '<option>Erro ao listar colunas</option>';
      setStatus($status, 'Falha ao buscar colunas (fallback). Último erro: ' + lastErr, true);
    }

    // Quando troca de board (team)
    $board.addEventListener('change', function () {
      const teamId = this.value;
      const t = boardsCache.find(b => b.id === teamId);
      loadColumnsForTeam(teamId, t ? t.name : teamId);
    });

    // --- ATUALIZAR WORK ITEM ---
    // Tenta usar WorkItemTracking RestClient; se não, faz PATCH via fetch
    function updateWorkItemState(id, newState, comment) {
      return new Promise(function (resolve, reject) {
        // Primeiro: tentar usar WIT RestClient
        VSS.require(['TFS/WorkItemTracking/RestClient'], function (WitRestClient) {
          try {
            const witClient = WitRestClient.getClient();
            const patch = [
              { op: 'add', path: '/fields/System.History', value: comment || 'Movido via extensão' },
              { op: 'add', path: '/fields/System.State', value: newState }
            ];
            witClient.updateWorkItem(patch, id).then(function (updated) {
              resolve(updated);
            }).catch(function (err) {
              warn('WitRestClient.updateWorkItem falhou, fallback PATCH fetch:', err);
              // fallback para fetch PATCH
              fallbackPatchWorkItem(id, patch).then(resolve).catch(reject);
            });
          } catch (e) {
            warn('WitRestClient error fallback', e);
            fallbackPatchWorkItem(id, [
              { op: 'add', path: '/fields/System.History', value: comment || 'Movido via extensão' },
              { op: 'add', path: '/fields/System.State', value: newState }
            ]).then(resolve).catch(reject);
          }
        }, function (requireErr) {
          warn('WIT RestClient não disponível, fallback PATCH:', requireErr);
          fallbackPatchWorkItem(id, [
            { op: 'add', path: '/fields/System.History', value: comment || 'Movido via extensão' },
            { op: 'add', path: '/fields/System.State', value: newState }
          ]).then(resolve).catch(reject);
        });
      });
    }

    // Fallback PATCH via REST
    async function fallbackPatchWorkItem(id, patchOps) {
      const url = collectionBase(webContext) + encodeURIComponent(webContext.project.name) + '/_apis/wit/workitems/' + id + '?api-version=5.1';
      log('fallbackPatchWorkItem', url, patchOps);
      const res = await fetchWithCredentials(url, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json-patch+json' },
        body: JSON.stringify(patchOps)
      });
      if (!res.ok) {
        const txt = await res.text().catch(()=>'');
        throw new Error('PATCH falhou: ' + res.status + ' - ' + txt);
      }
      return res.json();
    }

    // Clique no botão mover
    $moveBtn.addEventListener('click', async function () {
      setStatus($status, 'Executando mover...', false);
      $moveBtn.disabled = true;
      try {
        if (!workItemId) await ensureWorkItemId();
        if (!workItemId) {
          setStatus($status, 'Work item não identificado. Abra um work item e tente novamente.', true);
          $moveBtn.disabled = false;
          return;
        }

        const selectedColumnValue = $column.value;
        const selectedColumnIndex = $column.selectedIndex;
        const selectedColumnText = $column.options[selectedColumnIndex] ? $column.options[selectedColumnIndex].text : selectedColumnValue;

        // encontrar objeto de coluna
        let columnObj = columnsCache.find(c => (c.id && String(c.id) === String(selectedColumnValue)) || (c.name && c.name === selectedColumnText));
        if (!columnObj) columnObj = columnsCache.find(c => c.name === selectedColumnText) || null;

        // tentar extrair mapping para state
        let targetState = null;
        if (columnObj) {
          if (columnObj.mappedStates && columnObj.mappedStates.length) {
            const ms = columnObj.mappedStates[0];
            targetState = typeof ms === 'string' ? ms : (ms.name || ms.value || null);
          } else if (columnObj.states && columnObj.states.length) {
            const s = columnObj.states[0];
            targetState = typeof s === 'string' ? s : (s.name || s.value || null);
          } else if (columnObj.stateMappings && columnObj.stateMappings.length) {
            const s = columnObj.stateMappings[0];
            targetState = typeof s === 'string' ? s : (s.state || s.name || null);
          }
        }

        if (targetState) {
          setStatus($status, `Atualizando State -> ${targetState} ...`, false);
          await updateWorkItemState(workItemId, targetState, `Movido via extensão para ${selectedColumnText}`);
          setStatus($status, `Work item ${workItemId} atualizado. State = "${targetState}"`, false);
          try {
            const wiService = await VSS.getService(VSS.ServiceIds.WorkItemFormService);
            if (wiService && wiService.refresh) wiService.refresh();
          } catch (e) { /* não crítico */ }
          $moveBtn.disabled = false;
          return;
        }

        // se não há mapping de state, informar usuário (ou tentar API de boards se desejar)
        setStatus($status, 'Não foi possível mapear coluna para State automaticamente. Se desejar, posso tentar mover via Boards API (requer endpoints suportados).', true);
      } catch (err) {
        error('Erro mover:', err);
        setStatus($status, 'Erro ao mover: ' + (err && err.message ? err.message : err), true);
      } finally {
        $moveBtn.disabled = false;
      }
    });

    // Inicial: carregar boards
    loadBoards();
  }

  // --- Registro da contribuição ---
  VSS.ready(function () {
    try {
      const contributionId = 'JeffersonMota.move-card-onprem.move-card-form-group';
      VSS.register(contributionId, function (context) {
        log('contribution handler invoked', context && context.instanceContext ? context.instanceContext.host : null);
        const root = document.getElementById('move-card-root');
        if (!root) {
          error('Elemento root (#move-card-root) não encontrado.');
          VSS.notifyLoadFailed('Elemento root não encontrado');
          return {};
        }
        if (!root._controllerCreated) {
          root._controllerCreated = true;
          createController(root);
        }
        return {
          onContextUpdate: function (updatedContext) {
            log('onContextUpdate', updatedContext);
            // Se quiser reagir a mudanças de contexto (ex: outro work item aberto), podemos recarregar aqui.
          }
        };
      });

      VSS.notifyLoadSucceeded();
      log('registered and notified loaded');
    } catch (err) {
      error('Erro ao registrar contribution:', err);
      try { VSS.notifyLoadFailed(err && err.message ? err.message : 'Erro desconhecido'); } catch (_) {}
    }
  });

})();
