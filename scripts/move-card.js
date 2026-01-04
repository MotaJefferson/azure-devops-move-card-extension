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
      $statusEl.classList.toggle('error', Boolean(isError));
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
    let workItemType = null;
    let teamsCache = [];
    let columnsCache = [];
    const selectedBoards = new Map();

    // Obtém workItemId e tipo via WorkItemFormService
    async function ensureWorkItemInfo() {
      if (workItemId && workItemType) return { workItemId, workItemType };
      try {
        const wiService = await VSS.getService(VSS.ServiceIds.WorkItemFormService);
        workItemId = await wiService.getId();
        workItemType = await wiService.getFieldValue('System.WorkItemType');
        log('workItem', { workItemId, workItemType });
        return { workItemId, workItemType };
      } catch (err) {
        throw new Error('Erro ao obter dados do work item: ' + err);
      }
    }

    function apiBase() {
      return collectionBase(webContext);
    }

    function buildUrl(path) {
      return apiBase() + path.replace(/^\/+/, '');
    }

    // --- LISTAR TIMES ---
    async function loadBoards() {
      $board.disabled = true;
      $board.innerHTML = '<option>Carregando...</option>';
      setStatus($status, 'Listando times...', false);

      const projectName = encodeURIComponent(webContext.project.name);
      const projectId = encodeURIComponent(webContext.project.id);

      const urls = [
        buildUrl(`_apis/projects/${projectId}/teams?api-version=5.0`),
        buildUrl(`_apis/projects/${projectName}/teams?api-version=5.0`),
        buildUrl(`${projectName}/_apis/teams?api-version=5.0`),
        buildUrl(`_apis/teams?api-version=5.0`)
      ];

      let lastErr = null;
      for (const url of urls) {
        try {
          const res = await fetchWithCredentials(url);
          if (!res.ok) {
            const txt = await res.text().catch(() => '');
            lastErr = `HTTP ${res.status} ${res.statusText} - ${txt}`;
            warn('loadBoards non-ok', url, lastErr);
            continue;
          }
          const json = await res.json();
          const teams = json.value || [];
          if (!teams.length) {
            lastErr = 'Nenhum time retornado de ' + url;
            continue;
          }
          teamsCache = teams.map(t => ({ id: t.id, name: t.name || t.displayName || t.id }));
          populateBoardSelect();
          setStatus($status, 'Times carregados', false);
          if (teamsCache.length) loadColumnsForTeam(teamsCache[0].id);
          return;
        } catch (err) {
          warn('loadBoards error', url, err);
          lastErr = err;
        }
      }

      $board.innerHTML = '<option>Erro</option>';
      setStatus($status, 'Erro ao listar times. Último erro: ' + lastErr, true);
    }

    function populateBoardSelect() {
      $board.innerHTML = '';
      if (!teamsCache || !teamsCache.length) {
        $board.innerHTML = '<option>(nenhum time)</option>';
        $board.disabled = true;
        return;
      }
      teamsCache.forEach(team => {
        const opt = document.createElement('option');
        opt.value = team.id;
        opt.textContent = team.name;
        $board.appendChild(opt);
      });
      $board.disabled = false;
    }

    // --- LISTAR COLUNAS PARA UM TIME ---
    async function loadColumnsForTeam(teamId) {
      const team = teamsCache.find(t => t.id === teamId);
      $column.disabled = true;
      $column.innerHTML = '<option>Carregando...</option>';
      setStatus($status, 'Listando colunas para o time: ' + (team ? team.name : teamId), false);
      columnsCache = [];

      await ensureWorkItemInfo();

      const projectName = encodeURIComponent(webContext.project.name);
      const teamSegment = team ? encodeURIComponent(team.name) : encodeURIComponent(teamId);

      // Descobrir board do time
      const boardListUrls = [
        buildUrl(`${projectName}/${teamSegment}/_apis/work/boards?api-version=5.0`),
        buildUrl(`${projectName}/${teamId}/_apis/work/boards?api-version=5.0`),
        buildUrl(`_apis/work/teams/${teamId}/boards?api-version=5.0`)
      ];

      let board = selectedBoards.get(teamId) || null;
      let lastErr = null;

      if (!board) {
        for (const url of boardListUrls) {
          try {
            const res = await fetchWithCredentials(url);
            if (!res.ok) {
              const txt = await res.text().catch(() => '');
              lastErr = `HTTP ${res.status} ${res.statusText} - ${txt}`;
              warn('board list non-ok', url, lastErr);
              continue;
            }
            const json = await res.json();
            const boards = json.value || [];
            if (!boards.length) {
              lastErr = 'Nenhum board retornado de ' + url;
              continue;
            }
            board = boards.find(b => b.isDefault) || boards[0];
            selectedBoards.set(teamId, board);
            break;
          } catch (err) {
            warn('board list error', url, err);
            lastErr = err;
          }
        }
      }

      if (!board) {
        $column.innerHTML = '<option>Erro ao achar board</option>';
        setStatus($status, 'Não foi possível identificar um board para este time. Último erro: ' + lastErr, true);
        return;
      }

      // Buscar colunas do board
      const columnUrls = [
        buildUrl(`${projectName}/${teamSegment}/_apis/work/boards/${board.id}/columns?api-version=5.0`),
        buildUrl(`${projectName}/${teamId}/_apis/work/boards/${board.id}/columns?api-version=5.0`),
        buildUrl(`_apis/work/teams/${teamId}/boards/${board.id}/columns?api-version=5.0`)
      ];

      lastErr = null;
      for (const url of columnUrls) {
        try {
          const res = await fetchWithCredentials(url);
          if (!res.ok) {
            const txt = await res.text().catch(() => '');
            lastErr = `HTTP ${res.status} ${res.statusText} - ${txt}`;
            warn('columns non-ok', url, lastErr);
            continue;
          }
          const json = await res.json();
          const cols = json.value || json.columns || [];
          if (!cols.length) {
            lastErr = 'Nenhuma coluna retornada de ' + url;
            continue;
          }
          columnsCache = cols;
          populateColumnSelect();
          setStatus($status, 'Colunas carregadas', false);
          return;
        } catch (err) {
          warn('columns error', url, err);
          lastErr = err;
        }
      }

      $column.innerHTML = '<option>Erro</option>';
      setStatus($status, 'Falha ao buscar colunas. Último erro: ' + lastErr, true);
    }

    function mappedStateForColumn(column) {
      if (!column) return null;
      // VSTS/ADO columns normalmente expõem stateMappings com o tipo do work item
      if (column.stateMappings) {
        if (workItemType && column.stateMappings[workItemType]) {
          return column.stateMappings[workItemType];
        }
        const values = Object.keys(column.stateMappings).map(k => column.stateMappings[k]).filter(Boolean);
        if (values.length) return values[0];
      }
      if (column.mappedStates && column.mappedStates.length) {
        return column.mappedStates[0];
      }
      return null;
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
        const mappedState = mappedStateForColumn(c);
        opt.value = c.id || c.name || JSON.stringify(c);
        opt.textContent = mappedState ? `${c.name} • ${mappedState}` : (c.name || c.displayName || c.id || 'Coluna');
        if (mappedState) opt.dataset.state = mappedState;
        $column.appendChild(opt);
      });
      $column.disabled = false;
    }

    // Quando troca de board (team)
    $board.addEventListener('change', function () {
      const teamId = this.value;
      loadColumnsForTeam(teamId);
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
        const txt = await res.text().catch(() => '');
        throw new Error('PATCH falhou: ' + res.status + ' - ' + txt);
      }
      return res.json();
    }

    // Clique no botão mover
    $moveBtn.addEventListener('click', async function () {
      setStatus($status, 'Executando mover...', false);
      $moveBtn.disabled = true;
      try {
        const { workItemId: id } = await ensureWorkItemInfo();
        if (!id) {
          setStatus($status, 'Work item não identificado. Abra um work item e tente novamente.', true);
          $moveBtn.disabled = false;
          return;
        }

        const selectedColumnValue = $column.value;
        const selectedColumnIndex = $column.selectedIndex;
        const selectedColumnText = $column.options[selectedColumnIndex] ? $column.options[selectedColumnIndex].text : selectedColumnValue;
        const selectedColumnState = $column.options[selectedColumnIndex] ? $column.options[selectedColumnIndex].dataset.state : null;

        // encontrar objeto de coluna
        let columnObj = columnsCache.find(c => (c.id && String(c.id) === String(selectedColumnValue)) || (c.name && c.name === selectedColumnText));
        if (!columnObj) columnObj = columnsCache.find(c => c.name === selectedColumnText) || null;

        // tentar extrair mapping para state
        const targetState = selectedColumnState || mappedStateForColumn(columnObj);

        if (targetState) {
          setStatus($status, `Atualizando State -> ${targetState} ...`, false);
          await updateWorkItemState(id, targetState, `Movido via extensão para ${selectedColumnText}`);
          setStatus($status, `Work item ${id} atualizado. State = "${targetState}"`, false);
          try {
            const wiService = await VSS.getService(VSS.ServiceIds.WorkItemFormService);
            if (wiService && wiService.refresh) wiService.refresh();
          } catch (e) { /* não crítico */ }
          $moveBtn.disabled = false;
          return;
        }

        setStatus($status, 'Não foi possível mapear coluna para State automaticamente.', true);
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
