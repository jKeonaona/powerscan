'use strict';

function escHtml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function buildRowHtml(li) {
    const byBadge = li.proposed_by === 'skippy'
        ? '<span class="badge bg-info-subtle text-info-emphasis border border-info-subtle">skippy</span>'
        : '<span class="badge bg-secondary-subtle text-secondary-emphasis border border-secondary-subtle">user</span>';
    return `<tr data-line-item-id="${li.id}">
        <td class="text-muted small">&mdash;</td>
        <td><input type="text" class="form-control form-control-sm" data-field="element" aria-label="Element" value="${escHtml(li.element || '')}"></td>
        <td><input type="number" class="form-control form-control-sm" data-field="qty" aria-label="Qty" value="${li.qty ?? ''}" step="any"></td>
        <td><input type="number" class="form-control form-control-sm" data-field="length_ft" aria-label="Length (ft)" value="${li.length_ft ?? ''}" step="any"></td>
        <td><input type="number" class="form-control form-control-sm" data-field="height_ft" aria-label="Height (ft)" value="${li.height_ft ?? ''}" step="any"></td>
        <td><input type="text" class="form-control form-control-sm" data-field="dwg_ref" aria-label="Dwg Ref" value="${escHtml(li.dwg_ref || '')}"></td>
        <td><textarea class="form-control form-control-sm" data-field="notes" aria-label="Notes" rows="1">${escHtml(li.notes || '')}</textarea></td>
        <td>${byBadge}</td>
        <td class="text-center"><input type="checkbox" class="form-check-input" data-field="accepted" aria-label="Accepted" ${li.accepted ? 'checked' : ''}></td>
        <td class="text-nowrap">
            <button type="button" class="btn btn-sm btn-outline-primary save-btn me-1">Save</button>
            <button type="button" class="btn btn-sm btn-outline-danger delete-btn">Delete</button>
        </td>
    </tr>`;
}

function rowPayload(tr) {
    const payload = {};
    tr.querySelectorAll('[data-field]').forEach(el => {
        const field = el.dataset.field;
        if (el.type === 'checkbox') {
            payload[field] = el.checked;
        } else if (el.type === 'number') {
            payload[field] = el.value === '' ? null : parseFloat(el.value);
        } else {
            payload[field] = el.value.trim() || null;
        }
    });
    return payload;
}

function showRowError(tr, msg) {
    let errEl = tr.querySelector('.row-error');
    if (!errEl) {
        errEl = document.createElement('div');
        errEl.className = 'row-error text-danger small mt-1';
        tr.querySelector('td:last-child').appendChild(errEl);
    }
    errEl.textContent = msg;
}

function clearRowError(tr) {
    const errEl = tr.querySelector('.row-error');
    if (errEl) errEl.remove();
}

async function saveLineItem(buttonEl) {
    const tr = buttonEl.closest('tr');
    const id = tr.dataset.lineItemId;
    const payload = rowPayload(tr);
    buttonEl.disabled = true;
    try {
        const resp = await fetch(`/methodology-line-items/${id}`, {
            method: 'PATCH',
            credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            showRowError(tr, err.error || `Save failed (${resp.status})`);
            return;
        }
        clearRowError(tr);
        tr.classList.add('table-success');
        setTimeout(() => tr.classList.remove('table-success'), 1500);
    } catch (e) {
        showRowError(tr, 'Network error — save failed');
    } finally {
        buttonEl.disabled = false;
    }
}

async function deleteLineItem(buttonEl) {
    if (!window.confirm('Delete this line item? This cannot be undone.')) return;
    const tr = buttonEl.closest('tr');
    const id = tr.dataset.lineItemId;
    buttonEl.disabled = true;
    try {
        const resp = await fetch(`/methodology-line-items/${id}`, {
            method: 'DELETE',
            credentials: 'same-origin',
        });
        if (!resp.ok) {
            showRowError(tr, `Delete failed (${resp.status})`);
            buttonEl.disabled = false;
            return;
        }
        tr.remove();
    } catch (e) {
        showRowError(tr, 'Network error — delete failed');
        buttonEl.disabled = false;
    }
}

async function addLineItem() {
    const table = document.getElementById('line-items-table');
    const takeoffId = table.dataset.takeoffId;
    const errEl = document.getElementById('add-row-error');
    errEl.classList.add('d-none');

    const payload = {
        element: document.getElementById('new-element').value.trim() || null,
        qty: document.getElementById('new-qty').value === '' ? null : parseFloat(document.getElementById('new-qty').value),
        length_ft: document.getElementById('new-length-ft').value === '' ? null : parseFloat(document.getElementById('new-length-ft').value),
        height_ft: document.getElementById('new-height-ft').value === '' ? null : parseFloat(document.getElementById('new-height-ft').value),
        dwg_ref: document.getElementById('new-dwg-ref').value.trim() || null,
        notes: document.getElementById('new-notes').value.trim() || null,
    };

    const btn = document.getElementById('save-new-btn');
    btn.disabled = true;
    try {
        const resp = await fetch(`/methodology-takeoffs/${takeoffId}/line-items`, {
            method: 'POST',
            credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            errEl.textContent = err.error || `Add failed (${resp.status})`;
            errEl.classList.remove('d-none');
            return;
        }
        const li = await resp.json();
        const noRow = document.getElementById('no-items-row');
        if (noRow) noRow.remove();
        document.getElementById('line-items-tbody').insertAdjacentHTML('beforeend', buildRowHtml(li));
        ['new-element', 'new-qty', 'new-length-ft', 'new-height-ft', 'new-dwg-ref', 'new-notes'].forEach(id => {
            document.getElementById(id).value = '';
        });
        document.getElementById('add-row-form').classList.add('d-none');
        document.getElementById('add-row-btn').classList.remove('d-none');
    } catch (e) {
        errEl.textContent = 'Network error — add failed';
        errEl.classList.remove('d-none');
    } finally {
        btn.disabled = false;
    }
}

async function toggleAccepted(checkboxEl) {
    const lineItemId = checkboxEl.closest('tr').dataset.lineItemId;
    const isAccepted = checkboxEl.checked;

    try {
        const response = await fetch(`/methodology-line-items/${lineItemId}`, {
            method: 'PATCH',
            credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ accepted: isAccepted }),
        });

        if (!response.ok) {
            checkboxEl.checked = !isAccepted;
            alert(`Failed to ${isAccepted ? 'accept' : 'unaccept'} this row. Please refresh and try again.`);
            return;
        }

        const row = checkboxEl.closest('tr');
        if (row) {
            row.style.transition = 'background-color 0.3s';
            row.style.backgroundColor = '#d4edda';
            setTimeout(() => { row.style.backgroundColor = ''; }, 600);
        }
    } catch (err) {
        checkboxEl.checked = !isAccepted;
        alert(`Network error toggling acceptance: ${err.message}`);
    }
}

async function runStep2(buttonEl) {
    const takeoffId = buttonEl.dataset.takeoffId;
    const statusEl = document.getElementById('step2-status');
    buttonEl.disabled = true;
    statusEl.className = 'small mt-2 d-none';

    try {
        const resp = await fetch(`/methodology-takeoffs/${takeoffId}/propose-step-2`, {
            method: 'POST',
            credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({}),
        });
        const data = await resp.json().catch(() => ({}));
        if (resp.ok) {
            statusEl.textContent = data.content || 'Step 2 complete.';
            statusEl.className = 'small mt-2 text-success';
            setTimeout(() => window.location.reload(), 1500);
        } else {
            statusEl.textContent = data.error || `Error (${resp.status})`;
            statusEl.className = 'small mt-2 text-danger';
            buttonEl.disabled = false;
        }
    } catch (e) {
        statusEl.textContent = 'Network error — Step 2 failed.';
        statusEl.className = 'small mt-2 text-danger';
        buttonEl.disabled = false;
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // Event delegation on tbody — covers both server-rendered and dynamically added rows
    const tbody = document.getElementById('line-items-tbody');
    if (tbody) {
        tbody.addEventListener('click', e => {
            if (e.target.classList.contains('save-btn')) {
                saveLineItem(e.target);
            } else if (e.target.classList.contains('delete-btn')) {
                deleteLineItem(e.target);
            }
        });
    }

    const addRowBtn = document.getElementById('add-row-btn');
    const addRowForm = document.getElementById('add-row-form');
    const cancelAddBtn = document.getElementById('cancel-add-btn');
    const saveNewBtn = document.getElementById('save-new-btn');

    if (addRowBtn) {
        addRowBtn.addEventListener('click', () => {
            addRowForm.classList.remove('d-none');
            addRowBtn.classList.add('d-none');
            document.getElementById('new-element').focus();
        });
    }
    if (cancelAddBtn) {
        cancelAddBtn.addEventListener('click', () => {
            addRowForm.classList.add('d-none');
            addRowBtn.classList.remove('d-none');
        });
    }
    if (saveNewBtn) {
        saveNewBtn.addEventListener('click', addLineItem);
    }

    // Step 2 delegation (save/delete on step2 rows)
    const step2Tbody = document.getElementById('step2-tbody');
    if (step2Tbody) {
        step2Tbody.addEventListener('click', e => {
            if (e.target.classList.contains('save-btn')) {
                saveLineItem(e.target);
            } else if (e.target.classList.contains('delete-btn')) {
                deleteLineItem(e.target);
            }
        });
    }

    const runStep2Btn = document.getElementById('run-step2-btn');
    if (runStep2Btn) {
        runStep2Btn.addEventListener('click', () => runStep2(runStep2Btn));
    }

    // Auto-save Accepted checkbox toggles — no Save click required
    document.querySelectorAll('input[type="checkbox"][data-field="accepted"]').forEach(checkbox => {
        checkbox.addEventListener('change', () => toggleAccepted(checkbox));
    });
});
