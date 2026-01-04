/* globals VSS */
/*
  move-card.js - controlador do formulário de mover card (on-prem)
  - Usa RestClients do SDK (Core, Work, WorkItemTracking) para evitar CORS
  - Interface com combos alinhados ao cabeçalho do work item (headerItemPicker)
  - Atualiza System.State via WorkItemTracking RestClient
*/

(function () {
  'use strict';

  VSS.init({ explicitNotifyLoaded: true, usePlatformScripts: true });

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

  // --- Controller principal (cria UI e lógica) ---
  function createController(root, clients, Controls, Combos) {
    const $board = root.querySelector('#boardSelect');
    const $column = root.querySelector('#columnSelect');
    const $moveBtn = root.querySelector('#moveBtn');
    const $status = root.querySelector('#status');

    const boardPicker = Combos && Controls ? Controls.create(Combos.Combo, $board, {
      allowEdit: false,
      source: [],
      cssClass: 'headerItemPicker'
    }) : null;
    const columnPicker = Combos && Controls ? Controls.create(Combos.Combo, $column, {
      allowEdit: false,
      source: [],
      cssClass: 'headerItemPicker'
    }) : null;

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

    // --- LISTAR TIMES ---
    async function loadBoards() {
      setStatus($status, 'Listando times...', false);
      setBoardOptions([{ id: '', name: 'Carregando...' }], true);

      let lastErr = null;
      try {
        const teams = await clients.coreClient.getTeams(webContext.project.id, true, 200);
        teamsCache = (teams || []).map(t => ({ id: t.id, name: t.name || t.displayName || t.id }));
        if (!teamsCache.length) {
          throw new Error('Nenhum time retornado pelo CoreRestClient');
        }
        setBoardOptions(teamsCache, false);
        setStatus($status, 'Times carregados', false);
        if (teamsCache.length) loadColumnsForTeam(teamsCache[0].id);
        return;
      } catch (err) {
        warn('loadBoards via SDK falhou', err);
        lastErr = err;
      }

      setBoardOptions([{ id: '', name: 'Erro' }], true);
      setStatus($status, 'Erro ao listar times. Último erro: ' + lastErr, true);
    }

    function setBoardOptions(teams, disabled) {
      if (boardPicker) {
        boardPicker.setSource(teams.map(t => ({ text: t.name, value: t.id })));
        boardPicker.setSelectedIndex(0);
        boardPicker.setEnabled(!disabled);
      } else {
        $board.innerHTML = '';
        teams.forEach(team => {
          const opt = document.createElement('option');
          opt.value = team.id;
          opt.textContent = team.name;
          $board.appendChild(opt);
        });
        $board.disabled = disabled;
      }
    }

    // --- LISTAR COLUNAS PARA UM TIME ---
    async function loadColumnsForTeam(teamId) {
      const team = teamsCache.find(t => t.id === teamId);
      setColumnOptions([{ name: 'Carregando...', id: '' }], true);
      setStatus($status, 'Listando colunas para o time: ' + (team ? team.name : teamId), false);
      columnsCache = [];

      await ensureWorkItemInfo();

      let board = selectedBoards.get(teamId) || null;
      let lastErr = null;

      if (!board) {
        try {
          const boards = await clients.workClient.getBoards({ projectId: webContext.project.id, teamId });
          if (boards && boards.length) {
            board = boards.find(b => b.isDefault) || boards[0];
            selectedBoards.set(teamId, board);
          } else {
            lastErr = 'Nenhum board retornado pelo WorkRestClient';
          }
        } catch (err) {
          warn('board list via SDK falhou', err);
          lastErr = err;
        }
      }

      if (!board) {
        setColumnOptions([{ id: '', name: 'Erro ao achar board' }], true);
        setStatus($status, 'Não foi possível identificar um board para este time. Último erro: ' + lastErr, true);
        return;
      }

      try {
        const cols = await clients.workClient.getBoardColumns({ projectId: webContext.project.id, teamId }, board.id);
        columnsCache = cols || [];
        if (!columnsCache.length) {
          throw new Error('Nenhuma coluna retornada pelo WorkRestClient');
        }
        setColumnOptions(columnsCache, false);
        setStatus($status, 'Colunas carregadas', false);
        return;
      } catch (err) {
        warn('columns via SDK falharam', err);
        lastErr = err;
      }

      setColumnOptions([{ id: '', name: 'Erro' }], true);
      setStatus($status, 'Falha ao buscar colunas. Último erro: ' + lastErr, true);
    }

    function mappedStateForColumn(column) {
      if (!column) return null;
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

    function columnLabel(c) {
      const mappedState = mappedStateForColumn(c);
      return mappedState ? `${c.name} • ${mappedState}` : (c.name || c.displayName || c.id || 'Coluna');
    }

    function setColumnOptions(columns, disabled) {
      if (columnPicker) {
        const items = columns.map(c => ({ text: columnLabel(c), value: c.id || c.name, column: c }));
        columnPicker.setSource(items);
        columnPicker.setSelectedIndex(0);
        columnPicker.setEnabled(!disabled);
      } else {
        $column.innerHTML = '';
        columns.forEach(c => {
          const opt = document.createElement('option');
          const mappedState = mappedStateForColumn(c);
          opt.value = c.id || c.name || JSON.stringify(c);
          opt.textContent = columnLabel(c);
          if (mappedState) opt.dataset.state = mappedState;
          $column.appendChild(opt);
        });
        $column.disabled = disabled;
      }
    }

    function getSelectedTeamId() {
      if (boardPicker) {
        const item = boardPicker.getSelectedItem();
        return item ? item.value : null;
      }
      return $board.value;
    }

    function getSelectedColumnInfo() {
      if (columnPicker) {
        const item = columnPicker.getSelectedItem();
        return item ? { value: item.value, text: item.text, column: item.column } : { value: null, text: null, column: null };
      }
      const selectedColumnIndex = $column.selectedIndex;
      const opt = $column.options[selectedColumnIndex];
      return {
        value: opt ? opt.value : null,
        text: opt ? opt.text : null,
        column: columnsCache.find(c => (c.id && String(c.id) === String(opt && opt.value)) || (c.name && c.name === (opt && opt.text))) || null,
        state: opt ? opt.dataset.state : null
      };
    }

    if (boardPicker) {
      boardPicker.getElement().on('change', function () {
        const teamId = getSelectedTeamId();
        loadColumnsForTeam(teamId);
      });
    } else {
      $board.addEventListener('change', function () {
        const teamId = this.value;
        loadColumnsForTeam(teamId);
      });
    }

    // --- ATUALIZAR WORK ITEM ---
    function updateWorkItemState(id, newState, comment) {
      const patch = [
        { op: 'add', path: '/fields/System.History', value: comment || 'Movido via extensão' },
        { op: 'add', path: '/fields/System.State', value: newState }
      ];
      return clients.witClient.updateWorkItem(patch, id);
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

        const selectedColumnInfo = getSelectedColumnInfo();
        const columnObj = selectedColumnInfo.column;
        const targetState = mappedStateForColumn(columnObj);
        const columnText = selectedColumnInfo.text || selectedColumnInfo.value;

        if (targetState) {
          setStatus($status, `Atualizando State -> ${targetState} ...`, false);
          await updateWorkItemState(id, targetState, `Movido via extensão para ${columnText}`);
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

  function loadModules() {
    return new Promise((resolve, reject) => {
      VSS.require([
        'TFS/Core/RestClient',
        'TFS/Work/RestClient',
        'TFS/WorkItemTracking/RestClient',
        'VSS/Controls',
        'VSS/Controls/Combos'
      ], function (CoreRestClient, WorkRestClient, WitRestClient, Controls, Combos) {
          try {
            const clients = {
              coreClient: CoreRestClient.getClient(),
              workClient: WorkRestClient.getClient(),
              witClient: WitRestClient.getClient()
            };
            resolve({ clients, Controls, Combos });
          } catch (e) {
            reject(e);
          }
        }, reject);
    });
  }

  // --- Registro da contribuição ---
  VSS.ready(async function () {
    try {
      const { clients, Controls, Combos } = await loadModules();
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
          createController(root, clients, Controls, Combos);
        }
        return {
          onContextUpdate: function (updatedContext) {
            log('onContextUpdate', updatedContext);
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
