document.addEventListener('DOMContentLoaded', function () {
	const form = document.getElementById('woodAppForm');
	let result = document.getElementById('result');
	const addMaterialBtn = document.getElementById('addMaterial');
	const materialsTableElem = document.getElementById('materialsTable');
	const materialsTable = materialsTableElem ? materialsTableElem.querySelector('tbody') : null;
	const resetBtn = document.getElementById('resetBtn');

	// Survey date default to today if empty
	const surveyInput = document.querySelector('input[name="survey_date"]');
	if (surveyInput && !surveyInput.value) {
		const today = new Date().toISOString().slice(0,10);
		surveyInput.value = today;
	}


	// --- Manufacturers loader and helpers ---
	let manufacturersList = [];

	async function loadManufacturers() {
		// Prefer fetching the external `stovemanufacturers.json` when available (served over HTTP).
		// Fall back to inline JSON only if the fetch fails (useful for file:// testing).
		try {
			const resp = await fetch('stovemanufacturers.json');
			if (resp && resp.ok) {
				const json = await resp.json();
				manufacturersList = Array.isArray(json) ? json : [];
				populateAllMakeControls();
				return manufacturersList;
			}
		} catch (e) {
			// fetch failed (possibly file:// or network); try inline next
			console.debug('stovemanufacturers.json fetch failed, will try inline data if present', e && e.message);
		}

		// fallback to inline JSON (works with file:// and no server)
		try {
			const inline = document.getElementById('manufacturers-data');
			if (inline) {
				const json = JSON.parse(inline.textContent || inline.innerText || '[]');
				manufacturersList = Array.isArray(json) ? json : [];
				populateAllMakeControls();
				return manufacturersList;
			}
		} catch (e) {
			console.debug('Failed to parse inline manufacturers-data', e && e.message);
		}

		// final fallback: empty list
		manufacturersList = [];
		populateAllMakeControls();
		return manufacturersList;
	}

	function populateAllMakeControls() {
		// For every row in the materials table, populate its make control according to its type
		if (!materialsTable) return;
		Array.from(materialsTable.querySelectorAll('tr')).forEach(row => {
			populateMakeForRow(row);
		});
	}

	// load manufacturers in background
	loadManufacturers();

	async function initializeManufacturerCombo(input, typeVal = '') {
		if (!input) return;
		const list = await loadManufacturersForType(typeVal);
		
		// Remove any existing native datalist to avoid conflicts
		if (input.getAttribute('list')) input.removeAttribute('list');

		// Ensure we have a wrapper for the combo control
		let wrapper = input.closest('.combo-wrapper');
		if (!wrapper) {
			wrapper = document.createElement('div');
			wrapper.className = 'combo-wrapper';
			input.parentNode.insertBefore(wrapper, input);
			wrapper.appendChild(input);
			// toggle button to show all
			const btn = document.createElement('button');
			btn.type = 'button';
			btn.className = 'combo-toggle';
			btn.setAttribute('aria-label', 'Show manufacturers');
			btn.textContent = '▾';
			wrapper.appendChild(btn);
			// list element
			const listEl = document.createElement('ul');
			listEl.className = 'combo-list';
			wrapper.appendChild(listEl);

			// event wiring
			let currentItems = [];
			function renderList(filter) {
				const ll = listEl;
				ll.innerHTML = '';
				const f = (filter || '').toString().toLowerCase();
				const filtered = currentItems.filter(i => i.toString().toLowerCase().includes(f));
				filtered.slice(0,200).forEach((item, idx) => {
					const li = document.createElement('li');
					li.className = 'combo-item';
					li.textContent = item;
					li.tabIndex = -1;
					li.dataset.index = idx;
					li.addEventListener('click', function () {
						input.value = item;
						input.dispatchEvent(new Event('input', { bubbles: true }));
						closeList();
					});
					ll.appendChild(li);
				});
				// if no results, show a disabled item
				if (!filtered.length) {
					const li = document.createElement('li');
					li.className = 'combo-item';
					li.textContent = 'No matches';
					li.style.opacity = '0.6';
					ll.appendChild(li);
				}
				// reset active index
				wrapper._activeIndex = -1;
			}

			function openList() { wrapper.classList.add('open'); }
			function closeList() { wrapper.classList.remove('open'); wrapper._activeIndex = -1; updateActive(); }
			function toggleList() { if (wrapper.classList.contains('open')) closeList(); else { renderList(input.value || ''); openList(); } }

			function updateActive() {
				const items = Array.from(listEl.querySelectorAll('.combo-item'));
				items.forEach((it,i) => it.classList.toggle('active', i === (wrapper._activeIndex || -1)));
				if (wrapper._activeIndex >= 0 && items[wrapper._activeIndex]) {
					const el = items[wrapper._activeIndex];
					// ensure visible
					const rect = el.getBoundingClientRect();
					const parentRect = listEl.getBoundingClientRect();
					if (rect.top < parentRect.top) el.scrollIntoView(true);
					else if (rect.bottom > parentRect.bottom) el.scrollIntoView(false);
				}
			}

			// keyboard nav
			input.addEventListener('keydown', function (ev) {
				const items = Array.from(listEl.querySelectorAll('.combo-item'));
				if (ev.key === 'ArrowDown') {
					ev.preventDefault();
					if (!wrapper.classList.contains('open')) { renderList(input.value || ''); openList(); }
					wrapper._activeIndex = Math.min((wrapper._activeIndex || -1) + 1, items.length - 1);
					updateActive();
				} else if (ev.key === 'ArrowUp') {
					ev.preventDefault();
					wrapper._activeIndex = Math.max((wrapper._activeIndex || -1) - 1, 0);
					updateActive();
				} else if (ev.key === 'Enter') {
					if (wrapper.classList.contains('open') && (wrapper._activeIndex || -1) >= 0) {
						ev.preventDefault();
						const sel = items[wrapper._activeIndex];
						if (sel) {
							input.value = sel.textContent;
							input.dispatchEvent(new Event('input', { bubbles: true }));
						}
						closeList();
					}
				} else if (ev.key === 'Escape') {
					if (wrapper.classList.contains('open')) { ev.preventDefault(); closeList(); }
				}
			});

			input.addEventListener('input', function () { renderList(input.value || ''); if (!wrapper.classList.contains('open')) openList(); });
			btn.addEventListener('click', function (e) { e.preventDefault(); toggleList(); });

			// close on outside click
			document.addEventListener('click', function (ev) {
				if (!wrapper.contains(ev.target)) closeList();
			});

			// store a method to update the available items from outside
			wrapper._setItems = function (items) { currentItems = Array.isArray(items) ? items : []; renderList(input.value || ''); };
		}

		// populate the wrapper's list using the manufacturers list
		try {
			if (wrapper && wrapper._setItems) wrapper._setItems(list || []);
		} catch (e) {
			// graceful fallback: if anything fails, leave input as-is
			console.debug('Failed to initialize combobox for manufacturer', e && e.message);
		}
	}

	function slugifyType(v) {
		if (!v) return '';
		return v.toString().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
	}

	async function loadManufacturersForType(typeValue) {
		const key = slugifyType(typeValue || '');
		// try inline per-type JSON
		try {
			if (key) {
				const inline = document.getElementById('manufacturers-data-' + key);
				if (inline) {
					const json = JSON.parse(inline.textContent || inline.innerText || '[]');
					return Array.isArray(json) ? json : [];
				}
			}
		} catch (e) {}

		// try external per-type file
		if (key) {
			try {
				const resp = await fetch(`manufacturers-${key}.json`);
				if (resp && resp.ok) {
					const json = await resp.json();
					return Array.isArray(json) ? json : [];
				}
			} catch (e) {}
		}

		// fallback to global list (inline or fetched)
		if (manufacturersList && manufacturersList.length) return manufacturersList;
		// attempt to load the global list synchronously if not already loaded
		await loadManufacturers();
		return manufacturersList;
	}

	// populate the Make select for a specific table row based on its Type
	async function populateMakeForRow(row) {
		if (!row) return;
		const typeSel = row.querySelector('select[name="type"]') || row.querySelector('#type');
		let makeEl = row.querySelector('select[name="make"]') || row.querySelector('input[name="make"]');
		if (!makeEl) return;
		// prefer the actual option value; avoid falling back to option text to prevent unexpected fetches
		const typeVal = typeSel ? (typeSel.value || '') : '';
		
		if (makeEl.tagName.toLowerCase() === 'select') {
			// populate select
			const list = await loadManufacturersForType(typeVal);
			makeEl.innerHTML = '';
			const placeholder = document.createElement('option');
			placeholder.value = '';
			placeholder.textContent = 'Select';
			makeEl.appendChild(placeholder);
			list.forEach(m => {
				const opt = document.createElement('option');
				opt.value = m;
				opt.textContent = m;
				makeEl.appendChild(opt);
			});
		} else {
			// Use the shared combo initialization function
			await initializeManufacturerCombo(makeEl, typeVal);
		}
	}

	// Persist completed_by selection
	const completedBySelect = document.querySelector('select[name="completed_by"]');
	try {
		if (completedBySelect) {
			const saved = localStorage.getItem('woodapp_completed_by');
			if (saved) completedBySelect.value = saved;
			completedBySelect.addEventListener('change', function () { localStorage.setItem('woodapp_completed_by', completedBySelect.value); });
		}
	} catch (e) { /* ignore storage errors */ }

	// Theme toggle (dark/light)
	const themeToggle = document.getElementById('themeToggle');
	(function initTheme() {
		try {
			const t = localStorage.getItem('woodapp_theme') || 'light';
			if (t === 'dark') document.body.classList.add('dark-mode');
		} catch (e) {}
	})();
	if (themeToggle) {
		themeToggle.addEventListener('click', function () {
			document.body.classList.toggle('dark-mode');
			try { localStorage.setItem('woodapp_theme', document.body.classList.contains('dark-mode') ? 'dark' : 'light'); } catch (e) {}
		});
	}

	// Display service worker cache version
	(function displayVersion() {
		const versionDisplay = document.getElementById('versionDisplay');
		if (versionDisplay && 'serviceWorker' in navigator) {
			navigator.serviceWorker.getRegistration().then(reg => {
				if (reg && reg.active) {
					// Fetch the service worker script to get the cache name
					fetch('./sw.js')
						.then(response => response.text())
						.then(text => {
							const match = text.match(/CACHE_NAME\s*=\s*['"]([^'"]+)['"]/);
							if (match) {
								versionDisplay.textContent = `v${match[1].replace('woodapp-v', '')}`;
							}
						})
						.catch(() => {});
				}
			}).catch(() => {});
		}
	})();

	// Clearance row evaluation: color rows based on required vs actual values and shielding
	function parseNumberForComparison(val) {
		if (val === null || val === undefined) return NaN;
		if (typeof val === 'number') return val;
		const s = String(val).trim();
		if (s === '') return NaN;
		// strip non-numeric except dot and minus
		const cleaned = s.replace(/[^0-9\.\-]/g, '');
		const n = parseFloat(cleaned);
		return Number.isFinite(n) ? n : NaN;
	}

	function evaluateClearanceRow(row) {
		if (!row) return;
		// ignore header rows (colspan)
		const first = row.querySelector('td');
		if (!first) return;
		if (first.getAttribute('colspan')) {
			row.classList.remove('row-positive','row-negative','row-caution');
			return;
		}
		const cells = row.querySelectorAll('td');
		const reqInput = cells[1] ? cells[1].querySelector('input, select, textarea') : null;
		const actInput = cells[2] ? cells[2].querySelector('input, select, textarea') : null;
		const reqVal = reqInput ? reqInput.value : (cells[1] ? cells[1].textContent.trim() : '');
		const actVal = actInput ? actInput.value : (cells[2] ? cells[2].textContent.trim() : '');
		const reqNum = parseNumberForComparison(reqVal);
		const actNum = parseNumberForComparison(actVal);
		// find shielded checkbox if present
		let shielded = false;
		if (cells.length >= 4) {
			const shieldCell = cells[3];
			if (shieldCell) {
				const chk = shieldCell.querySelector('input[type="checkbox"]');
				if (chk) shielded = !!chk.checked;
			}
		}
		
		row.classList.remove('row-positive','row-negative','row-caution');
		if (!Number.isNaN(reqNum) && !Number.isNaN(actNum)) {
			if (actNum >= reqNum) {
				row.classList.add('row-positive');
			} else {
				if (shielded) row.classList.add('row-caution');
				else row.classList.add('row-negative');
		}
	}
}

	// Attach listener to clearancesTable for input/change events (delegation)
	const clearancesTable = document.getElementById('clearancesTable');
	if (clearancesTable) {
		clearancesTable.addEventListener('input', function (e) {
			const row = e.target.closest('tr');
			if (row) evaluateClearanceRow(row);
		});
		clearancesTable.addEventListener('change', function (e) {
			const row = e.target.closest('tr');
			if (row) evaluateClearanceRow(row);
		});
		// initial evaluation
		Array.from(clearancesTable.querySelectorAll('tbody tr')).forEach(r => evaluateClearanceRow(r));
	}

	// we will not create or use a visible result area — PDF save will be offered on submit

	// Only wire add/remove if the buttons/table exist
	if (addMaterialBtn && materialsTable) {
		addMaterialBtn.addEventListener('click', function () {
			const tr = document.createElement('tr');
			tr.innerHTML = `
					<td data-label="Type"><input type="text" name="type"></td>
					<td data-label="Make"><select name="make"><option value="">Select</option></select></td>
					<td data-label="Model"><input type="text" name="model"></td>
					<td data-label="Installed By"><input type="text" name="installed_by"></td>
					<td data-label="Chimney Code"><input type="text" name="chimney_code"></td>
					<td data-label="Own/Shared"><input type="text" name="own_shared"></td>
					<td data-label="Chimney Condition"><input type="text" name="chimney_condition"></td>
					<td data-label="Label"><select name="label"><option value="">Select</option><option value="CSA B366">CSA B366</option><option value="ULC S610">ULC S610</option><option value="ULC S627">ULC S627</option><option value="ULC S628">ULC S628</option><option value="Other">Other</option><option value="Not Labelled">Not labelled</option></select></td>
					<td data-label="Location"><select name="location"><option value="">Select</option><option value="Dwelling">Dwelling</option><option value="Garage attached/detached">Garage attached/detached</option><option value="Outbuilding">Outbuilding</option></select></td>
					<td data-label="Actions"><button type="button" class="remove-row">×</button></td>
			`;
			materialsTable.appendChild(tr);
			// if there is a template type select in the DOM (the first row), clone its options into the new row
			const templateType = document.querySelector('#type');
			const typeInput = tr.querySelector('input[name="type"]');
			if (templateType && typeInput) {
				const newSel = document.createElement('select');
				newSel.name = 'type';
				newSel.innerHTML = templateType.innerHTML;
				typeInput.replaceWith(newSel);
			}

			// populate Make for this new row according to its Type (will use defaults if blank)
			populateMakeForRow(tr);
		});

		materialsTable.addEventListener('click', function (e) {
			if (e.target && e.target.classList.contains('remove-row')) {
				const row = e.target.closest('tr');
				if (row) row.remove();
			}
		});

		// when a Type select changes in the materials table, populate its Make select
		materialsTable.addEventListener('change', function (e) {
			const target = e.target;
			if (target && target.matches && target.matches('select[name="type"], select#type')) {
				const row = target.closest('tr');
				if (row) populateMakeForRow(row);
			}
		});

		// populate existing rows' Make selects based on their Type
		Array.from(materialsTable.querySelectorAll('tbody tr')).forEach(r => populateMakeForRow(r));
	}

	// collect rows generically (works for selects + inputs)
	function collectTableRows(tableBody) {
		if (!tableBody) return [];
		const rows = Array.from(tableBody.querySelectorAll('tr'));
		return rows.map(row => {
			const controls = Array.from(row.querySelectorAll('input, select, textarea'));
			const obj = {};
			controls.forEach((c, i) => {
				const key = c.name || c.id || `col${i}`;
				if (c.type === 'checkbox') obj[key] = !!c.checked;
				else obj[key] = c.value || '';
			});
			return obj;
		}).filter(o => Object.values(o).some(v => v !== '' && v !== false));
	}

	function collectFormData() {
		const fd = new FormData(form);
		const data = {};
		for (let pair of fd.entries()) {
			const [key, value] = pair;
			// handle array-style names
			if (key.endsWith('[]')) {
				const base = key.slice(0, -2);
				data[base] = data[base] || [];
				data[base].push(value);
			} else if (data[key] !== undefined) {
				// convert to array if multiple values
				if (!Array.isArray(data[key])) data[key] = [data[key]];
				data[key].push(value);
			} else {
				data[key] = value;
			}
		}

		// collect appliances/materials from the table if present
		data.appliances = collectTableRows(materialsTable);

		// photos if present
		const photosInput = document.getElementById('photos');
		if (photosInput && photosInput.files && photosInput.files.length) {
			data.photos = Array.from(photosInput.files).map(f => f.name);
		} else {
			data.photos = [];
		}

		// combine chimney_major and chimney_minor into chimney_code if present
		if (data.chimney_major !== undefined || data.chimney_minor !== undefined) {
			const maj = (data.chimney_major || '').toString();
			const min = (data.chimney_minor || '').toString();
			if (maj && min) data.chimney_code = `${maj}.${min}`;
			else if (maj) data.chimney_code = maj;
			// optional: remove the separate fields
			delete data.chimney_major;
			delete data.chimney_minor;
		}

		// Ensure shielded checkboxes (if present) are explicit booleans in the output
		const shieldInputs = document.querySelectorAll('input[name^="shielded_"]');
		if (shieldInputs && shieldInputs.length) {
			shieldInputs.forEach(inp => {
				data[inp.name] = !!inp.checked;
			});
		}

		return data;
	}

	// -- Saved entries (collected appliances + measurements) --
	const savedEntries = [];

	function updateSavedCount() {
		const el = document.getElementById('savedCount');
		if (!el) return;
		if (savedEntries.length === 0) el.textContent = '';
		else el.textContent = `${savedEntries.length} saved`;
	}

	function collectClearancesData(tableArg) {
		const table = tableArg || clearancesTable;
		const rows = Array.from((table ? table.querySelectorAll('tbody tr') : []) || []);
		const hasShielded = rows.some(r => Array.from(r.querySelectorAll('td')).some(td => td.querySelector && td.querySelector('input[type="checkbox"]')));
		const out = rows.map(r => {
			// skip hidden rows
			if (r.style.display === 'none') return null;
			const first = r.querySelector('td');
			if (!first) return null;
			const colspan = first.getAttribute('colspan');
			const label = first.textContent.trim();
			
			// Capture row styling class for PDF coloring
			let rowClass = '';
			if (r.classList.contains('row-positive')) rowClass = 'positive';
			else if (r.classList.contains('row-negative')) rowClass = 'negative';
			else if (r.classList.contains('row-caution')) rowClass = 'caution';
			
			if (colspan && parseInt(colspan) > 1) return { label, required: '', actual: '', shielded: '', _isHeader: true, _rowClass: rowClass };
			const cells = r.querySelectorAll('td');
			const reqInput = cells[1] ? cells[1].querySelector('input, select, textarea') : null;
			const actInput = cells[2] ? cells[2].querySelector('input, select, textarea') : null;
			const req = reqInput ? (reqInput.value || '') : (cells[1] ? cells[1].textContent.trim() : '');
			const act = actInput ? (actInput.value || '') : (cells[2] ? cells[2].textContent.trim() : '');
			let shieldVal = '';
			if (hasShielded) {
				const shieldCell = cells[3];
				if (shieldCell) {
					const chk = shieldCell.querySelector('input[type="checkbox"]');
					shieldVal = chk ? (chk.checked ? 'Yes' : 'No') : (shieldCell.textContent || '').trim();
				}
			}
			return { label, required: req, actual: act, shielded: shieldVal, _isHeader: false, _rowClass: rowClass };
		}).filter(Boolean);
		return out;
	}

	// Helper functions for section-specific operations
	function applySectionTypeRules(section, val) {
		// Find the clearances table - it's inside the section with "Measurements & Clearances" heading
		let clearancesTable = null;
		const sections = section.querySelectorAll('section');
		for (const sec of sections) {
			const h2 = sec.querySelector('h2');
			if (h2 && h2.textContent.includes('Measurements')) {
				clearancesTable = sec.querySelector('table');
				break;
			}
		}
		if (!clearancesTable) return;
		
		const v = (val || '').toString().toLowerCase();
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		
		// Show all rows first
		rows.forEach(r => r.style.display = '');
		
		// Helper to remove facing rows
		const removeFacingRows = () => {
			rows.forEach(r => {
				const first = r.querySelector('td');
				if (!first) return;
				const text = first.textContent.trim().toLowerCase();
				if (text === 'left facing' || text === 'right facing') r.remove();
			});
		};
		
		// Helper to check if facing row exists
		const facingRowExists = (label) => {
			return Array.from(clearancesTable.querySelectorAll('tbody tr')).some(r => {
				const td = r.querySelector('td');
				return td && td.textContent.trim().toLowerCase() === label.toLowerCase();
			});
		};
		
		// Helper to add facing rows after right side
		const addFacingRowsAfterRightSide = () => {
			if (facingRowExists('left facing') || facingRowExists('right facing')) return;
			
			const allRows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
			let insertAfter = null;
			for (const r of allRows) {
				const td = r.querySelector('td');
				if (td && td.textContent.trim().toLowerCase() === 'right side') {
					insertAfter = r;
					break;
				}
			}
			
			const tbody = clearancesTable.querySelector('tbody');
			const createFacingRow = (label) => {
				const tr = document.createElement('tr');
				const tdLabel = document.createElement('td');
				tdLabel.textContent = label;
				const tdReq = document.createElement('td');
				const reqInput = document.createElement('input');
				reqInput.type = 'text';
				reqInput.name = `${label.toLowerCase().replace(/\s+/g,'_')}_required`;
				tdReq.appendChild(reqInput);
				const tdAct = document.createElement('td');
				const actInput = document.createElement('input');
				actInput.type = 'text';
				actInput.name = `${label.toLowerCase().replace(/\s+/g,'_')}_actual`;
				tdAct.appendChild(actInput);
				tr.appendChild(tdLabel);
				tr.appendChild(tdReq);
				tr.appendChild(tdAct);
				
				// if shielded column present, append shield checkbox cell
				const theadRow = clearancesTable.querySelector('thead tr');
				if (theadRow && Array.from(theadRow.children).some(th => th.textContent.trim() === 'Shielded')) {
					const tdShield = document.createElement('td');
					const chk = document.createElement('input');
					chk.type = 'checkbox';
					chk.name = `shielded_${label.toLowerCase().replace(/\s+/g,'_')}`;
					chk.className = 'shielded-checkbox';
					tdShield.appendChild(chk);
					tr.appendChild(tdShield);
				}
				
				return tr;
			};
			
			const leftTr = createFacingRow('Left facing');
			const rightTr = createFacingRow('Right facing');
			if (insertAfter && insertAfter.parentNode) {
				insertAfter.parentNode.insertBefore(leftTr, insertAfter.nextSibling);
				insertAfter.parentNode.insertBefore(rightTr, leftTr.nextSibling);
			} else {
				tbody.appendChild(leftTr);
				tbody.appendChild(rightTr);
			}
		};
		
		// Helper to rename flue to chimney
		const renameFlueToChimney = (shouldRename) => {
			const mapping = [
				['flue pipe back', 'Chimney back'],
				['flue pipe side', 'Chimney side'],
				['flue pipe ceiling', 'Chimney ceiling']
			];
			Array.from(clearancesTable.querySelectorAll('tbody tr')).forEach(r => {
				const first = r.querySelector('td');
				if (!first) return;
				const text = first.textContent.trim().toLowerCase();
				mapping.forEach(([from, to]) => {
					if (shouldRename) {
						if (text.includes(from)) {
							first.textContent = to;
						}
					} else {
						// revert if currently renamed
						if (text.includes(to.toLowerCase())) {
							first.textContent = from.charAt(0).toUpperCase() + from.slice(1);
						}
					}
				});
			});
		};
		
	// Remove facing rows and reset flue labels before applying new rules
	removeFacingRows();
	renameFlueToChimney(false);
	
	// First, show all rows and restore their default values
	Array.from(clearancesTable.querySelectorAll('tbody tr')).forEach(r => {
		r.style.display = '';
		Array.from(r.querySelectorAll('input, select, textarea')).forEach(el => {
			if (el.type === 'checkbox' || el.type === 'radio') {
				el.checked = false;
			} else if (el.defaultValue !== undefined && el.defaultValue !== '') {
				el.value = el.defaultValue;
			}
		});
	});
	
	const hideByKeywords = (keywords) => {
		Array.from(clearancesTable.querySelectorAll('tbody tr')).forEach(r => {
			const first = r.querySelector('td');
			if (!first) return;
			const text = first.textContent.trim().toLowerCase();
			if (keywords.some(k => text.includes(k))) {
				r.style.display = 'none';
				Array.from(r.querySelectorAll('input, select, textarea')).forEach(el => {
					if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
					else el.value = '';
				});
			}
		});
	};		switch (v) {
			case 'kitchen wood range':
				hideByKeywords(['plenum','mantel']);
				break;
			case 'insert':
				hideByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum']);
				addFacingRowsAfterRightSide();
				break;
			case 'furnace':
				hideByKeywords(['left corner','right corner','mantel','top']);
				break;
			case 'boiler':
				hideByKeywords(['plenum','mantel']);
				break;
			case 'factory built fireplace':
				hideByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum','floor pad rear']);
				addFacingRowsAfterRightSide();
				break;
			case 'pellet stove/insert':
				hideByKeywords(['plenum']);
				break;
			case 'hearth':
				hideByKeywords(['plenum']);
				break;
			case 'outdoor wood boiler':
				renameFlueToChimney(true);
				hideByKeywords(['plenum','mantel']);
				break;
			case 'masonry fireplace':
				hideByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum','floor pad rear']);
				addFacingRowsAfterRightSide();
				break;
			case 'stove':
				hideByKeywords(['plenum','mantel']);
				break;
		}
	}

	// Apply flue pipe type clearances (SW=18", DW=6")
	function applyFluePipeTypeClearances(section, fluePipeType) {
		let clearancesTable = null;
		const sections = section.querySelectorAll('section');
		for (const sec of sections) {
			const h2 = sec.querySelector('h2');
			if (h2 && h2.textContent.includes('Measurements')) {
				clearancesTable = sec.querySelector('table');
				break;
			}
		}
		if (!clearancesTable) return;
		
		// Base clearance values (always use base values, shielding is applied per-row during evaluation)
		const baseValue = fluePipeType === 'SW' ? 18 : fluePipeType === 'DW' ? 6 : null;
		const requiredValue = baseValue !== null ? baseValue.toString() : '';
		
		// Update flue pipe rows
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			const label = firstCell.textContent.trim().toLowerCase();
			
			if (label.includes('flue pipe back') || label.includes('flue pipe side') || label.includes('flue pipe ceiling')) {
				const requiredInput = row.querySelectorAll('td input')[0];
				if (requiredInput && requiredValue) {
					requiredInput.value = requiredValue;
					// Trigger evaluation
					evaluateClearanceRow(row);
				}
			}
		});
	}

	// Helper function to update required field value when shielded checkbox changes for flue pipe rows
	function updateFluePipeRequiredValue(row, checkbox) {
		const firstCell = row.querySelector('td');
		if (!firstCell) return;
		const label = firstCell.textContent.trim().toLowerCase();
		
		// Only apply to flue pipe rows
		if (!label.includes('flue pipe back') && !label.includes('flue pipe side') && !label.includes('flue pipe ceiling')) {
			return;
		}
		
		const cells = row.querySelectorAll('td');
		const requiredInput = cells[1] ? cells[1].querySelector('input') : null;
		if (!requiredInput) return;
		
		const currentValue = parseFloat(requiredInput.value);
		if (isNaN(currentValue)) return;
		
		if (checkbox.checked) {
			// Reduce by 50%
			requiredInput.value = (currentValue / 2).toString();
		} else {
			// Restore to 100% (double it)
			requiredInput.value = (currentValue * 2).toString();
		}
		
		// Trigger evaluation
		evaluateClearanceRow(row);
	}

	function addShieldedColumnToSection(section) {
		// Find the clearances table - it's inside the section with "Measurements & Clearances" heading
		let clearancesTable = null;
		const sections = section.querySelectorAll('section');
		for (const sec of sections) {
			const h2 = sec.querySelector('h2');
			if (h2 && h2.textContent.includes('Measurements')) {
				clearancesTable = sec.querySelector('table');
				break;
			}
		}
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		if (Array.from(theadRow.children).some(th => th.textContent.trim() === 'Shielded')) return;
		
		const th = document.createElement('th');
		th.textContent = 'Shielded';
		theadRow.appendChild(th);
		
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			const colspan = firstCell.getAttribute('colspan');
			if (colspan && parseInt(colspan) > 1) {
			firstCell.setAttribute('colspan', parseInt(colspan) + 1);
			return;
		}
		const td = document.createElement('td');
		const input = document.createElement('input');
		input.type = 'checkbox';
		input.name = `shielded_${Math.random().toString(36).substr(2, 9)}`;
		input.className = 'shielded-checkbox';
		
		// Add event listener to update required value for flue pipe rows
		input.addEventListener('change', function() {
			updateFluePipeRequiredValue(row, input);
		});
		
		td.appendChild(input);
		row.appendChild(td);
	});
}	function removeShieldedColumnFromSection(section) {
		// Find the clearances table - it's inside the section with "Measurements & Clearances" heading
		let clearancesTable = null;
		const sections = section.querySelectorAll('section');
		for (const sec of sections) {
			const h2 = sec.querySelector('h2');
			if (h2 && h2.textContent.includes('Measurements')) {
				clearancesTable = sec.querySelector('table');
				break;
			}
		}
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		const ths = Array.from(theadRow.children);
		let shieldIndex = -1;
		for (let i = 0; i < ths.length; i++) {
			if (ths[i].textContent.trim() === 'Shielded') { shieldIndex = i; break; }
		}
		if (shieldIndex >= 0) theadRow.removeChild(ths[shieldIndex]);
		
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			const colspan = firstCell.getAttribute('colspan');
			if (colspan && parseInt(colspan) > 1) {
				firstCell.setAttribute('colspan', Math.max(1, parseInt(colspan) - 1));
				return;
			}
			const cells = row.querySelectorAll('td');
			if (shieldIndex >= 0 && cells.length > shieldIndex) {
				const cell = cells[shieldIndex];
				if (cell && cell.querySelector('.shielded-checkbox')) cell.remove();
			}
		});
	}
	
	function createManufacturerCombobox(inputElement) {
		if (!inputElement) return;
		// Check if already wrapped (to avoid double-wrapping on clones)
		if (inputElement.parentNode && inputElement.parentNode.classList.contains('combobox-wrapper')) {
			return; // Already has combobox
		}
		// Use the loaded manufacturer list
		if (manufacturersList && manufacturersList.length > 0) {
			const wrapper = document.createElement('div');
			wrapper.className = 'combobox-wrapper';
			wrapper.style.position = 'relative';
			wrapper.style.width = '100%';
			
			inputElement.parentNode.insertBefore(wrapper, inputElement);
			wrapper.appendChild(inputElement);
			
			const datalist = document.createElement('datalist');
			datalist.id = 'manufacturers-' + Math.random().toString(36).substr(2, 9);
			manufacturersList.forEach(name => {
				const option = document.createElement('option');
				option.value = name;
				datalist.appendChild(option);
			});
			wrapper.appendChild(datalist);
			inputElement.setAttribute('list', datalist.id);
		}
	}

	function resetApplianceAndMeasurements() {
		// clear materials table: first row inputs cleared, other rows removed
		if (materialsTable) {
			const rows = Array.from(materialsTable.querySelectorAll('tr'));
			rows.forEach((r, i) => {
				if (i === 0) {
					Array.from(r.querySelectorAll('input, select, textarea')).forEach(inp => {
						if (inp.type === 'checkbox' || inp.type === 'radio') inp.checked = false;
						else if (inp.tagName.toLowerCase() === 'select') inp.selectedIndex = 0;
						else inp.value = '';
					});
				} else {
					r.remove();
				}
			});
			// repopulate make control for first row
			const firstRow = materialsTable.querySelector('tr');
			if (firstRow) populateMakeForRow(firstRow);
		}

		// clear clearances inputs to their default values (use defaultValue if present)
		if (clearancesTable) {
			Array.from(clearancesTable.querySelectorAll('input, select, textarea')).forEach(inp => {
				if (inp.type === 'checkbox' || inp.type === 'radio') inp.checked = false;
				else if (inp.tagName.toLowerCase() === 'select') inp.selectedIndex = 0;
				else if (typeof inp.defaultValue !== 'undefined') inp.value = inp.defaultValue || '';
				else inp.value = '';
			});
			// clear evaluation classes
			Array.from(clearancesTable.querySelectorAll('tbody tr')).forEach(r => r.classList.remove('row-positive','row-negative','row-caution'));
		}

		// clear remarks
		const remarks = form.querySelector('textarea[name="remarks"]');
		if (remarks) remarks.value = '';
	}

	// wire the Add Another Appliance button (if present)
	const addAnotherBtn = document.getElementById('addAnotherBtn');
	if (addAnotherBtn) {
		addAnotherBtn.addEventListener('click', async function () {
			try {
				const container = document.getElementById('appliancesContainer');
				const buttonsContainer = document.getElementById('buttonsContainer');
				if (!container || !buttonsContainer) return;
				
				// Get all existing appliance sections
				const sections = container.querySelectorAll('.appliance-section');
				const newNum = sections.length + 1;
				
				// Clone the first section as template
				const template = sections[0];
				const clone = template.cloneNode(true);
				clone.setAttribute('data-appliance-num', newNum);
				
				// Update the heading
				const heading = clone.querySelector('h2');
				if (heading && heading.textContent === 'Appliance') {
					heading.textContent = `Appliance ${newNum}`;
				}
				
				// Remove IDs from cloned elements to avoid duplicates
				clone.querySelectorAll('[id]').forEach(el => {
					el.removeAttribute('id');
				});
				
				// Clear all input values
				clone.querySelectorAll('input[type="text"], input[type="number"], input[type="email"], input[type="date"], textarea').forEach(inp => {
					if (inp.hasAttribute('value') && inp.getAttribute('value')) {
						// Keep default values like "18" and "8"
						inp.value = inp.getAttribute('value');
					} else {
						inp.value = '';
					}
				});
				
				clone.querySelectorAll('select').forEach(sel => {
					// Reset to first option or selected default
					const defaultSelected = sel.querySelector('option[selected]');
					if (defaultSelected) {
						sel.value = defaultSelected.value;
					} else {
						sel.selectedIndex = 0;
					}
				});
				
			clone.querySelectorAll('input[type="checkbox"], input[type="radio"]').forEach(inp => {
				inp.checked = false;
			});
			
			// Remove any shielded column from the cloned section (in case source had shielding=yes)
			removeShieldedColumnFromSection(clone);
			
			// Remove any styling classes from clearances rows
			clone.querySelectorAll('tbody tr').forEach(r => {
				r.classList.remove('row-positive','row-negative','row-caution');
				r.style.display = ''; // Make sure all rows are visible
			});				// Ensure tables maintain proper display structure
				clone.querySelectorAll('table').forEach(table => {
					table.style.display = '';
					table.style.width = '100%';
					table.style.borderCollapse = 'collapse';
				});
				
				// Ensure thead and tbody have proper display
				clone.querySelectorAll('thead').forEach(thead => {
					thead.style.display = '';
				});
				clone.querySelectorAll('tbody').forEach(tbody => {
					tbody.style.display = '';
				});
				
				// Insert the clone before the buttons container
				container.appendChild(clone);
				
				// Initialize event handlers for the new section
				const typeSelect = clone.querySelector('select[name="type"]');
				if (typeSelect) {
					typeSelect.addEventListener('change', function() {
						const clearancesTable = clone.querySelector('table');
						if (clearancesTable) {
							applySectionTypeRules(clone, typeSelect.value);
						}
					});
				}
				
			const shieldingSelect = clone.querySelector('select[name="shielding"]');
			if (shieldingSelect) {
			shieldingSelect.addEventListener('change', function() {
				const clearancesTable = clone.querySelector('table');
				if (clearancesTable) {
					const val = (shieldingSelect.value || '').toString().toLowerCase();
					if (val === 'yes') addShieldedColumnToSection(clone);
					else removeShieldedColumnFromSection(clone);
				}
			});
		}
		
		// Add listener for flue pipe type changes in cloned appliance
		const fluePipeTypeSelect = clone.querySelector('select[name="flue_pipe_type"]');
		if (fluePipeTypeSelect) {
			fluePipeTypeSelect.addEventListener('change', function() {
				applyFluePipeTypeClearances(clone, fluePipeTypeSelect.value);
			});
		}				// Populate manufacturer dropdown for new section
				const makeInput = clone.querySelector('input[name="make"]');
				if (makeInput) {
					// Remove the cloned combo-wrapper structure and rebuild it fresh
					let wrapper = makeInput.closest('.combo-wrapper');
					if (wrapper) {
						// Extract the input from the wrapper
						const parent = wrapper.parentNode;
						parent.insertBefore(makeInput, wrapper);
						wrapper.remove();
					}
			// Now reinitialize with the same logic as appliance 1
			await initializeManufacturerCombo(makeInput);
		}
		
		// Attach clearance row evaluation listeners to the cloned section
		const clonedClearancesTable = clone.querySelector('table:not(.materials-table)');
		if (clonedClearancesTable) {
			clonedClearancesTable.addEventListener('input', function (e) {
				const row = e.target.closest('tr');
				if (row) evaluateClearanceRow(row);
			});
			clonedClearancesTable.addEventListener('change', function (e) {
				const row = e.target.closest('tr');
				if (row) evaluateClearanceRow(row);
			});
			// Initial evaluation for all rows
			Array.from(clonedClearancesTable.querySelectorAll('tbody tr')).forEach(r => evaluateClearanceRow(r));
		}
		
		// Update delete button visibility and renumber
		updateApplianceSections();			// Scroll to new section
			clone.scrollIntoView({ behavior: 'smooth', block: 'start' });			} catch (e) {
				console.error('Failed to add appliance section', e);
				alert('Unable to add appliance: ' + (e && e.message ? e.message : e));
			}
		});
	}


// Update delete button visibility and renumber appliances
function updateApplianceSections() {
	const container = document.getElementById('appliancesContainer');
	if (!container) return;
	
	const sections = container.querySelectorAll('.appliance-section:not(.collapsing)');
	const hasMultiple = sections.length > 1;
	
	sections.forEach((section, index) => {
		const num = index + 1;
		section.setAttribute('data-appliance-num', num);
		
		const heading = section.querySelector('.appliance-heading');
		if (heading) {
			heading.textContent = num === 1 ? 'Appliance' : `Appliance ${num}`;
		}
		
		const deleteBtn = section.querySelector('.delete-appliance-btn');
		if (deleteBtn) {
			deleteBtn.style.display = hasMultiple ? 'block' : 'none';
		}
	});
}

// Delete appliance section with animation
function deleteApplianceSection(section) {
	const num = section.getAttribute('data-appliance-num');
	
	if (!confirm(`Are you sure you want to remove appliance ${num}?`)) {
		return;
	}
	
	// Add collapsing class to trigger animation
	section.classList.add('collapsing');
	
	// Remove from DOM after animation completes
	setTimeout(() => {
		section.remove();
		updateApplianceSections();
	}, 400); // Match CSS transition duration
}

// Attach delete handlers to existing and new appliance sections
function attachDeleteHandlers() {
	const container = document.getElementById('appliancesContainer');
	if (!container) return;
	
	container.addEventListener('click', function(e) {
		if (e.target.classList.contains('delete-appliance-btn')) {
			const section = e.target.closest('.appliance-section');
			if (section) {
				deleteApplianceSection(section);
			}
		}
	});
}

// Initialize delete button visibility on page load
attachDeleteHandlers();
updateApplianceSections();


	// ---------- PDF generation and save helpers ----------
	function sanitizeFilename(name) {
		return name.replace(/[\\/:*?"<>|]+/g, '').trim();
	}

	async function saveBlobWithPicker(blob, suggestedName) {
		// Try File System Access API first
		if (window.showSaveFilePicker) {
			const opts = {
				suggestedName,
				types: [
					{
						description: 'PDF',
						accept: { 'application/pdf': ['.pdf'] }
					}
				]
			};
			const handle = await window.showSaveFilePicker(opts);
			const writable = await handle.createWritable();
			await writable.write(blob);
			await writable.close();
			return;
		}

		// Fallback: trigger anchor download (browser may save to Downloads)
		const url = URL.createObjectURL(blob);
		const a = document.createElement('a');
		a.href = url;
		a.download = suggestedName;
		document.body.appendChild(a);
		a.click();
		a.remove();
		setTimeout(() => URL.revokeObjectURL(url), 5000);
	}

	// generatePdfBlob: create the PDF and return the blob + suggested filename
	async function generatePdfBlob(data) {
		const { jsPDF } = window.jspdf || {};
		if (!jsPDF) throw new Error('jsPDF library not loaded');
		const doc = new jsPDF({ unit: 'pt', format: 'a4' });
		const margin = 40;
		const pageWidth = doc.internal.pageSize.getWidth();
		let cursorY = margin;

		// (copy of the PDF generation logic from createAndSavePdf)
		doc.setFillColor(246, 250, 255);
		doc.rect(0, 0, pageWidth, 64, 'F');
		doc.setFontSize(18);
		doc.setFont('helvetica', 'bold');
		doc.setTextColor(22, 55, 92);
		doc.text('WOOD APP FORM', pageWidth / 2, cursorY, { align: 'center' });
		doc.setTextColor(0, 0, 0);
		cursorY += 26;

		const policy = data.policy || '';
		const surveyDate = data.survey_date || '';
		const completedBy = data.completed_by || '';

		doc.setFontSize(10);
		doc.setFont('helvetica', 'normal');
		doc.autoTable({
			startY: cursorY,
			head: [['Policy # or Name', 'Survey date', 'Completed by']],
			body: [[policy, surveyDate, completedBy]],
			theme: 'grid',
			styles: { fontSize: 10 },
			headStyles: { fillColor: [225, 235, 245], textColor: 22, halign: 'center' },
			columnStyles: { 0: { cellWidth: 120 }, 1: { cellWidth: 120 }, 2: { cellWidth: 120 } }
		});
		cursorY = doc.lastAutoTable.finalY + 12;

		let chimneyLegend = [];
		if (Array.isArray(data.appliances) && data.appliances.length) {
			const appHead = ['Type', 'Make', 'Model', 'Installed By', 'Chimney Code', 'Own/Shared', 'Chimney Condition', 'Shielding', 'Label', 'Location', 'Flue Pipe Type'];
			const appBody = data.appliances.map((app, idx) => {
				const maj = app.chimney_major || app['chimney_major'] || '';
				const min = app.chimney_minor || app['chimney_minor'] || '';
				let chimneyCode = app.chimney_code || app['chimney_code'] || '';
				if ((!chimneyCode || chimneyCode === '') && maj) {
					chimneyCode = min ? `${maj}.${min}` : `${maj}`;
				}
				let chimneyFullWords = '';
				try {
					if (materialsTable) {
						const rows = Array.from(materialsTable.querySelectorAll('tr'));
						const row = rows[idx];
						if (row) {
							const majSel = row.querySelector('select[name="chimney_major"]');
							const minSel = row.querySelector('select[name="chimney_minor"]');
							const majText = majSel && majSel.selectedOptions && majSel.selectedOptions[0] ? majSel.selectedOptions[0].text : '';
							const minText = minSel && minSel.selectedOptions && minSel.selectedOptions[0] ? minSel.selectedOptions[0].text : '';
							if (majText || minText) {
								const cleanMaj = majText ? majText.replace(/^\s*\d+\s*[—-]?\s*/,'').trim() : '';
								const cleanMin = minText ? minText.replace(/^\s*\d+\s*[\-]?\s*/,'').trim() : '';
								chimneyFullWords = [cleanMaj, cleanMin].filter(Boolean).join(' / ');
							}
						}
					}
				} catch (e) {}
			if (chimneyCode) chimneyLegend.push({ code: chimneyCode, words: chimneyFullWords });
			
			// Map flue_pipe_type to display text
			let fluePipeDisplay = '';
			if (app.flue_pipe_type === 'SW') {
				fluePipeDisplay = 'SW (18")';
			} else if (app.flue_pipe_type === 'DW') {
				fluePipeDisplay = 'DW (6")';
			} else {
				fluePipeDisplay = 'N/A';
			}
			
			return [
				app.type || app['type'] || (app['col0'] || ''),
				app.make || '',
				app.model || '',
				app.installed_by || app['installed_by'] || '',
				chimneyCode,
				app.own_shared || app['own_shared'] || '',
				app.chimney_condition || app['chimney_condition'] || '',
				(app.shielding === true || app.shielding === 'yes' || app.shielding === 'Yes') ? 'Yes' : (app.shielding === 'no' || app.shielding === false ? 'No' : (app.shielding || '')),
				app.label || '',
				app.location || '',
				fluePipeDisplay
			];
		});			doc.setFontSize(12);
			doc.setFont('helvetica', 'bold');
			doc.text('Appliance Details', margin, cursorY);
			cursorY += 8;

			doc.autoTable({
				startY: cursorY,
				head: [appHead],
				body: appBody,
				styles: { fontSize: 9 },
				headStyles: { fillColor: [235, 245, 255], textColor: 22 },
				theme: 'striped',
				columnStyles: { 0: { cellWidth: 60 }, 1: { cellWidth: 70 }, 2: { cellWidth: 70 }, 3: { cellWidth: 70 } }
			});
			cursorY = doc.lastAutoTable.finalY + 12;
		}

		// clearances and notes: reuse same logic as existing flow
		if (clearancesTable) {
			const hasShielded = Array.from(clearancesTable.querySelectorAll('thead th')).some(th => th.textContent.trim() === 'Shielded');
			const head = hasShielded ? ['Clearances from', 'Required', 'Actual', 'Shielded'] : ['Clearances from', 'Required', 'Actual'];
			const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
			const body = rows.map(r => {
				const first = r.querySelector('td');
				if (!first) return null;
				const colspan = first.getAttribute('colspan');
				const label = first.textContent.trim();
				if (colspan && parseInt(colspan) > 1) return { label: label, required: '', actual: '', shielded: '', _isHeader: true };
				const cells = r.querySelectorAll('td');
				const reqInput = cells[1] ? cells[1].querySelector('input, select, textarea') : null;
				const actInput = cells[2] ? cells[2].querySelector('input, select, textarea') : null;
				const req = reqInput ? (reqInput.value || '') : (cells[1] ? cells[1].textContent.trim() : '');
				const act = actInput ? (actInput.value || '') : (cells[2] ? cells[2].textContent.trim() : '');
				let shieldVal = '';
				if (hasShielded) {
					const shieldCell = cells[3];
					if (shieldCell) {
						const chk = shieldCell.querySelector('input[type="checkbox"]');
						shieldVal = chk ? (chk.checked ? 'Yes' : 'No') : shieldCell.textContent.trim();
					}
				}
				return { label, required: req, actual: act, shielded: shieldVal, _isHeader: false };
			}).filter(Boolean);

			if (body.length) {
				doc.setFontSize(12);
				doc.setFont('helvetica', 'bold');
				doc.text('Measurements & Clearances', margin, cursorY);
				cursorY += 8;
				const atBody = body.map(b => hasShielded ? [b.label, b.required, b.actual, b.shielded] : [b.label, b.required, b.actual]);
				doc.autoTable({
					startY: cursorY,
					head: [head],
					body: atBody,
					styles: { fontSize: 9 },
					headStyles: { fillColor: [235, 245, 255], textColor: 22 },
					theme: 'grid',
					didParseCell: function (dataCell) {
						const raw = dataCell.row && dataCell.row.raw;
						if (!raw) return;
						const isHeader = raw[1] === '' && raw[2] === '' && (hasShielded ? raw[3] === '' : true);
						if (isHeader) {
							if (dataCell.column.index === 0) {
								dataCell.cell.colSpan = hasShielded ? 4 : 3;
								// Only apply blue background to Floor Pad row
								const label = raw[0] || '';
								if (label.toLowerCase().includes('floor pad')) {
									dataCell.cell.styles.fillColor = [235, 245, 255];
									dataCell.cell.styles.textColor = 22;
								}
								dataCell.cell.styles.halign = 'left';
							} else {
								dataCell.cell.text = '';
							}
						} else {
							// Apply row coloring based on _rowClass (only to body rows, not header)
							if (dataCell.row.section === 'body') {
								const bodyItem = body[dataCell.row.index];
								if (bodyItem && bodyItem._rowClass) {
									if (bodyItem._rowClass === 'positive') {
										dataCell.cell.styles.fillColor = [230, 255, 230];
									} else if (bodyItem._rowClass === 'negative') {
										dataCell.cell.styles.fillColor = [255, 235, 235];
									} else if (bodyItem._rowClass === 'caution') {
										dataCell.cell.styles.fillColor = [255, 250, 230];
									}
								}
							}
						}
					}
				});
				cursorY = doc.lastAutoTable.finalY + 12;
			}
		}

		// Notes
		doc.setFontSize(12);
		doc.setFont('helvetica', 'bold');
		doc.text('Notes', margin, cursorY);
		cursorY += 12;
		doc.setFont('helvetica', 'normal');
		doc.setFontSize(10);
		const notesText = (data.remarks || '').trim();
		const notesLines = notesText ? doc.splitTextToSize(notesText, pageWidth - margin * 2) : [];
		if (notesLines.length) {
			doc.text(notesLines, margin, cursorY);
			cursorY += notesLines.length * 12 + 8;
		}

		if (Array.isArray(chimneyLegend) && chimneyLegend.length) {
			const uniq = {};
			chimneyLegend.forEach(item => { if (!item || !item.code) return; if (!uniq[item.code]) uniq[item.code] = item.words || ''; else if (!uniq[item.code] && item.words) uniq[item.code] = item.words; });
			const legendLines = [];
			for (const code of Object.keys(uniq)) {
				const words = uniq[code];
				if (words) legendLines.push(`${code} — ${words}`);
				else legendLines.push(`${code}`);
			}
			if (legendLines.length) {
				doc.setFont('helvetica', 'bold');
				doc.text('Chimney Code Legend', margin, cursorY);
				cursorY += 12;
				doc.setFont('helvetica', 'normal');
				const wrapped = doc.splitTextToSize(legendLines.join('\n'), pageWidth - margin * 2);
				doc.text(wrapped, margin, cursorY);
				cursorY += wrapped.length * 12 + 8;
			}
		}

		const blob = doc.output('blob');
		const safePolicy = sanitizeFilename(policy || 'policy');
		const safeDate = sanitizeFilename(surveyDate || (new Date()).toISOString().slice(0,10));
		const suggestedName = `${safePolicy} - ${safeDate} - Wood app form.pdf`;
		return { blob, suggestedName };
	}

	// modify createAndSavePdf to use the generator and then offer the picker
	async function createAndSavePdf(data) {
		const res = await generatePdfBlob(data);
		await saveBlobWithPicker(res.blob, res.suggestedName);
		return res;
	}

	// MSAL + Graph upload helper (small-file PUT)
	async function ensureMsalAvailable() {
		if (!ONEDRIVE_CLIENT_ID) return null;
		if (!window.msal || !window.msal.PublicClientApplication) throw new Error('MSAL not loaded');
		const msalConfig = { auth: { clientId: ONEDRIVE_CLIENT_ID, redirectUri: window.location.origin } };
		const msalInstance = new msal.PublicClientApplication(msalConfig);
		return msalInstance;
	}

	async function getAccessToken(msalInstance) {
		const scopes = ["Files.ReadWrite"];
		try {
			const accounts = msalInstance.getAllAccounts();
			if (accounts && accounts.length) {
				const silentReq = { account: accounts[0], scopes };
				const silent = await msalInstance.acquireTokenSilent(silentReq);
				return silent.accessToken;
			}
		} catch (e) {
			// fallback to interactive
		}
		// interactive login
		const loginResp = await msalInstance.loginPopup({ scopes });
		const tokenResp = await msalInstance.acquireTokenSilent({ account: loginResp.account, scopes }).catch(async () => {
			return await msalInstance.acquireTokenPopup({ scopes });
		});
		return tokenResp.accessToken;
	}

	async function uploadToOneDrive(blob, filename) {
		if (!ONEDRIVE_CLIENT_ID) throw new Error('ONEDRIVE_CLIENT_ID not configured');
		const msalInstance = await ensureMsalAvailable();
		if (!msalInstance) throw new Error('MSAL not available');
		const token = await getAccessToken(msalInstance);
		// small-file PUT to root
		const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(filename)}:/content`;
		const resp = await fetch(url, {
			method: 'PUT',
			headers: {
				'Authorization': `Bearer ${token}`,
				'Content-Type': 'application/pdf'
			},
			body: blob
		});
		if (!resp.ok) {
			const txt = await resp.text().catch(() => '');
			throw new Error(`Upload failed: ${resp.status} ${resp.statusText} ${txt}`);
		}
		return await resp.json();
	}

	// wire the Save-to-OneDrive button
	const saveOneDriveBtn = document.getElementById('saveOneDriveBtn');
	if (saveOneDriveBtn) {
		if (!ONEDRIVE_CLIENT_ID) {
			// hide or disable if not configured
			saveOneDriveBtn.style.display = 'none';
		} else {
			saveOneDriveBtn.addEventListener('click', async function () {
				try {
					const data = collectFormData();
					const { blob, suggestedName } = await generatePdfBlob(data);
					// upload
					saveOneDriveBtn.disabled = true;
					saveOneDriveBtn.textContent = 'Uploading...';
					await uploadToOneDrive(blob, suggestedName);
					alert('Saved to OneDrive as ' + suggestedName);
				} catch (err) {
					console.error('OneDrive upload failed', err);
					alert('OneDrive upload failed: ' + (err && err.message ? err.message : err));
				} finally {
					saveOneDriveBtn.disabled = false;
					saveOneDriveBtn.textContent = 'Save to OneDrive';
				}
			});
		}
	}

	// Share / Save button: mobile uses native share sheet, desktop downloads the PDF
	const shareBtn = document.getElementById('shareBtn');

	// Generate a PDF blob for multiple saved entries + optional current entry
	async function generatePdfBlobForAll(meta, entries) {
		const { jsPDF } = window.jspdf || {};
		if (!jsPDF) throw new Error('jsPDF library not loaded');
		const doc = new jsPDF({ unit: 'pt', format: 'a4' });
		const margin = 40;
		const pageWidth = doc.internal.pageSize.getWidth();
		let cursorY = margin;

			// We'll render one appliance per page. Each page shows the header/meta, a single-appliance table,
			// then that entry's clearances and notes immediately under it.
			const policy = meta.policy || '';
			const surveyDate = meta.survey_date || '';
			const completedBy = meta.completed_by || '';

			let firstPage = true;

			// compute total number of appliance pages to decide whether to label Notes with appliance index
			let totalAppliances = 0;
			entries.forEach(en => { if (en && Array.isArray(en.appliances) && en.appliances.length) totalAppliances += en.appliances.length; });
			let applianceCounter = 0;
			// track whether any per-entry notes were actually rendered; if none are rendered
			// but remarks exist, we'll append a final Notes page as a fallback so users never
			// lose their entered notes.
			let notesRenderedCount = 0;

			// helper to render the common header + meta table at the top of a page
			function renderPageHeader() {
				doc.setFillColor(246, 250, 255);
				doc.rect(0, 0, pageWidth, 64, 'F');
				doc.setFontSize(18);
				doc.setFont('helvetica', 'bold');
				doc.setTextColor(22, 55, 92);
				doc.text('WOOD APP FORM', pageWidth / 2, margin, { align: 'center' });
				doc.setTextColor(0, 0, 0);
				cursorY = margin + 26;
			}

			// ensure there's room for X points, otherwise add a page and re-render header
			function ensureSpace(pointsNeeded) {
				const pageHeight = doc.internal.pageSize.getHeight();
				if (cursorY + pointsNeeded > pageHeight - margin) {
					doc.addPage();
					renderPageHeader();
					return true;
				}
				return false;
			}

			// collect chimney legend entries (code -> wording) across appliances
			let chimneyLegend = [];

			for (let eIdx = 0; eIdx < entries.length; eIdx++) {
				const entry = entries[eIdx] || {};
				const appliances = Array.isArray(entry.appliances) && entry.appliances.length ? entry.appliances : [null];

				for (let aIdx = 0; aIdx < appliances.length; aIdx++) {
					const app = appliances[aIdx];
					applianceCounter += 1;

					if (!firstPage) {
						doc.addPage();
					}
					
					// header band + title
					doc.setFillColor(246, 250, 255);
					doc.rect(0, 0, pageWidth, 64, 'F');
					doc.setFontSize(18);
					doc.setFont('helvetica', 'bold');
					doc.setTextColor(22, 55, 92);
					doc.text('WOOD APP FORM', pageWidth / 2, margin, { align: 'center' });
					doc.setTextColor(0, 0, 0);
					cursorY = margin + 26;

					// Show policy table only on first page
					if (firstPage) {
						doc.setFontSize(10);
						doc.setFont('helvetica', 'normal');
						doc.autoTable({
							startY: cursorY,
							head: [['Policy # or Name', 'Survey date', 'Completed by']],
							body: [[policy, surveyDate, completedBy]],
							theme: 'grid',
							styles: { fontSize: 10 },
							headStyles: { fillColor: [225, 235, 245], textColor: 22, halign: 'center' },
							columnStyles: { 0: { cellWidth: 120 }, 1: { cellWidth: 120 }, 2: { cellWidth: 120 } }
						});
						cursorY = doc.lastAutoTable.finalY + 12;
					}
					
					firstPage = false;

					// Appliance details rendered as a wrapped key/value table to match the entry layout
					if (app) {
						const maj = app.chimney_major || app['chimney_major'] || '';
						const min = app.chimney_minor || app['chimney_minor'] || '';
						let chimneyCode = app.chimney_code || '';
						if ((!chimneyCode || chimneyCode === '') && maj) chimneyCode = min ? `${maj}.${min}` : `${maj}`;
						// determine human-friendly wording for this chimney code
						let chimneyFullWords = app.chimney_full_words || '';
						if (!chimneyFullWords) {
							try {
								const majOpt = document.querySelector('#chimney_major option[value="' + maj + '"]');
								const minOpt = document.querySelector('#chimney_minor option[value="' + min + '"]');
								const majText = majOpt ? majOpt.textContent : '';
								const minText = minOpt ? minOpt.textContent : '';
								if (majText || minText) {
									const cleanMaj = majText ? majText.replace(/^\s*\d+\s*[—-]?\s*/,'').trim() : '';
									const cleanMin = minText ? minText.replace(/^\s*\d+\s*[\-]?\s*/,'').trim() : '';
									chimneyFullWords = [cleanMaj, cleanMin].filter(Boolean).join(' / ');
								}
							} catch (e) {}
						}
						const infoRows = [
							['Type', app.type || ''],
							['Manufacturer', app.make || ''],
							['Model', app.model || ''],
							['Installed By', app.installed_by || ''],
							['Chimney Code', chimneyCode || ''],
							['Own/Shared', app.own_shared || ''],
							['Chimney Condition', app.chimney_condition || ''],
							['Shielding', (app.shielding === true || app.shielding === 'yes' || app.shielding === 'Yes') ? 'Yes' : (app.shielding || '')],
							['Label', app.label || ''],
							['Location', app.location || ''],
						['Flue Pipe Type', (app.flue_pipe_type === 'SW' ? 'SW (18")' : (app.flue_pipe_type === 'DW' ? 'DW (6")' : 'N/A'))]
					];
						doc.setFontSize(12);
						doc.setFont('helvetica', 'bold');
						doc.text('Appliance Details', margin, cursorY);
						cursorY += 8;
						doc.autoTable({
							startY: cursorY,
							head: [['Field', 'Value']],
							body: infoRows,
							styles: { fontSize: 10 },
							headStyles: { fillColor: [235, 245, 255], textColor: 22 },
							theme: 'grid',
							columnStyles: { 0: { cellWidth: 140 }, 1: { cellWidth: pageWidth - margin * 2 - 140 } },
							didParseCell: function (data) {
								if (data.section === 'body' && data.column.index === 0) {
									data.cell.styles.fontStyle = 'bold';
									data.cell.styles.fillColor = [245, 250, 255];
								}
							}
						});
						cursorY = doc.lastAutoTable.finalY + 10;
					} else {
						doc.setFontSize(12);
						doc.setFont('helvetica', 'bold');
						doc.text('Appliance Details', margin, cursorY);
						cursorY += 18;
						doc.setFontSize(10);
						doc.text('(No appliance details provided)', margin, cursorY);
						cursorY += 18;
					}

					// Clearances for this entry
					if (Array.isArray(entry.clearances) && entry.clearances.length) {
						const hasShielded = entry.clearances.some(r => r.shielded !== undefined && r.shielded !== '');
						const head = hasShielded ? ['Clearances from', 'Required', 'Actual', 'Shielded'] : ['Clearances from', 'Required', 'Actual'];
						const atBody = entry.clearances.map(b => hasShielded ? [b.label, b.required, b.actual, b.shielded] : [b.label, b.required, b.actual]);
						doc.setFontSize(12);
						doc.setFont('helvetica', 'bold');
						doc.text('Measurements & Clearances', margin, cursorY);
						cursorY += 8;
						doc.autoTable({
							startY: cursorY,
							head: [head],
							body: atBody,
							styles: { fontSize: 9 },
							headStyles: { fillColor: [235, 245, 255], textColor: 22 },
							theme: 'grid',
							didParseCell: function (dataCell) {
								const raw = dataCell.row && dataCell.row.raw;
								if (!raw) return;
								const isHeader = raw[1] === '' && raw[2] === '' && (hasShielded ? raw[3] === '' : true);
								if (isHeader) {
									if (dataCell.column.index === 0) {
										dataCell.cell.colSpan = hasShielded ? 4 : 3;
										// Only apply blue background to Floor Pad row
										const label = raw[0] || '';
										if (label.toLowerCase().includes('floor pad')) {
											dataCell.cell.styles.fillColor = [235, 245, 255];
											dataCell.cell.styles.textColor = 22;
										}
										dataCell.cell.styles.halign = 'left';
									} else {
										dataCell.cell.text = '';
									}
								} else {
									// Apply row coloring based on _rowClass (only to body rows, not header)
									if (dataCell.row.section === 'body') {
										const bodyItem = entry.clearances[dataCell.row.index];
										if (bodyItem && bodyItem._rowClass) {
										if (bodyItem._rowClass === 'positive') {
											dataCell.cell.styles.fillColor = [230, 255, 230];
										} else if (bodyItem._rowClass === 'negative') {
											dataCell.cell.styles.fillColor = [255, 235, 235];
										} else if (bodyItem._rowClass === 'caution') {
											dataCell.cell.styles.fillColor = [255, 250, 230];
										}
									}
								}
							}
						}
					});
						cursorY = doc.lastAutoTable.finalY + 10;
					}

					// Notes for this entry (placed directly under the appliance/clearances)
					if (entry.remarks && entry.remarks.trim()) {
						doc.setFontSize(12);
						doc.setFont('helvetica', 'bold');
						const notesTitle = totalAppliances > 1 ? `Notes — Appliance ${applianceCounter}` : 'Notes';
						// ensure there's room for the notes title + a couple lines; if not, add a page and re-render header
						ensureSpace(48);
						doc.text(notesTitle, margin, cursorY);
						cursorY += 12;
						doc.setFont('helvetica', 'normal');
						doc.setFontSize(10);
						const notesText = entry.remarks.trim();
						const notesLines = notesText ? doc.splitTextToSize(notesText, pageWidth - margin * 2) : [];
						if (notesLines.length) {
							// ensure enough space for the notes; if not, add page and re-render header
							if (cursorY + notesLines.length * 12 > doc.internal.pageSize.getHeight() - margin) {
								doc.addPage();
								renderPageHeader();
							}
							doc.text(notesLines, margin, cursorY);
							cursorY += notesLines.length * 12 + 8;
							// mark that we actually rendered notes for at least one entry
							notesRenderedCount++;
						}
					}

					// Chimney Code Legend for this appliance (placed directly under notes)
					if (app) {
						const maj = app.chimney_major || app['chimney_major'] || '';
						const min = app.chimney_minor || app['chimney_minor'] || '';
						let chimneyCode = app.chimney_code || '';
						if ((!chimneyCode || chimneyCode === '') && maj) chimneyCode = min ? `${maj}.${min}` : `${maj}`;
						
						if (chimneyCode) {
							let chimneyFullWords = app.chimney_full_words || '';
							if (!chimneyFullWords) {
								try {
									const majOpt = document.querySelector('#chimney_major option[value="' + maj + '"]');
									const minOpt = document.querySelector('#chimney_minor option[value="' + min + '"]');
									const majText = majOpt ? majOpt.textContent : '';
									const minText = minOpt ? minOpt.textContent : '';
									if (majText || minText) {
										const cleanMaj = majText ? majText.replace(/^\s*\d+\s*[—-]?\s*/,'').trim() : '';
										const cleanMin = minText ? minText.replace(/^\s*\d+\s*[\-]?\s*/,'').trim() : '';
										chimneyFullWords = [cleanMaj, cleanMin].filter(Boolean).join(' / ');
									}
								} catch (e) {}
							}
							
							if (chimneyFullWords) {
								ensureSpace(36);
								doc.setFont('helvetica', 'bold');
								doc.setFontSize(10);
								doc.text('Chimney Code Legend', margin, cursorY);
								cursorY += 12;
								doc.setFont('helvetica', 'normal');
								const legendText = `${chimneyCode} — ${chimneyFullWords}`;
								const wrapped = doc.splitTextToSize(legendText, pageWidth - margin * 2);
								doc.text(wrapped, margin, cursorY);
								cursorY += wrapped.length * 12 + 8;
							}
						}
					}
				}
			}

			// If we didn't render any per-entry notes but there are remarks present
			// somewhere in the entries, append a final Notes page containing the
			// combined remarks so the user can find them in the PDF.
			if (notesRenderedCount === 0) {
				// collect any non-empty remarks
				const combined = entries.map(en => (en && en.remarks) ? en.remarks.trim() : '').filter(Boolean).join('\n\n');
				if (combined) {
					doc.addPage();
					renderPageHeader();
					doc.setFontSize(12);
					doc.setFont('helvetica', 'bold');
					doc.text('Notes', margin, cursorY);
					cursorY += 12;
					doc.setFont('helvetica', 'normal');
					doc.setFontSize(10);
					const lines = doc.splitTextToSize(combined, pageWidth - margin * 2);
					doc.text(lines, margin, cursorY);
					cursorY += lines.length * 12 + 8;
				}
			}

			const blob = doc.output('blob');
			const safePolicy = sanitizeFilename(policy || 'policy');
			const safeDate = sanitizeFilename(surveyDate || (new Date()).toISOString().slice(0,10));
			const suggestedName = `${safePolicy} - ${safeDate} - Wood app form.pdf`;
			return { blob, suggestedName };
	}
	
	// Collect data from all appliance sections
	function collectAllSections() {
		const container = document.getElementById('appliancesContainer');
		if (!container) return [];
		
		const sections = container.querySelectorAll('.appliance-section');
		const entries = [];
		
		sections.forEach(section => {
			const materialsTables = section.querySelectorAll('.materials-table');
			const clearancesTable = section.querySelector('table:not(.materials-table)');
			const remarksField = section.querySelector('textarea[name="remarks"]');
			
			// Each section represents ONE appliance entry
			const app = {};
			
			// Get data from first materials table (type, make, model, installed_by, chimney)
			if (materialsTables.length >= 1) {
				const firstTable = materialsTables[0].querySelector('tbody tr');
				if (firstTable) {
					const typeEl = firstTable.querySelector('select[name="type"]');
					if (typeEl) app.type = typeEl.value || '';
					const makeEl = firstTable.querySelector('input[name="make"]');
					if (makeEl) app.make = makeEl.value || '';
					const modelEl = firstTable.querySelector('input[name="model"]');
					if (modelEl) app.model = modelEl.value || '';
					const installedByEl = firstTable.querySelector('input[name="installed_by"]');
					if (installedByEl) app.installed_by = installedByEl.value || '';
					
					const chimMajEl = firstTable.querySelector('select[name="chimney_major"]');
					const chimMinEl = firstTable.querySelector('select[name="chimney_minor"]');
					if (chimMajEl) app.chimney_major = chimMajEl.value || '';
					if (chimMinEl) app.chimney_minor = chimMinEl.value || '';
					
					// Capture chimney words
					if (chimMajEl && chimMinEl) {
						try {
							const majText = chimMajEl.selectedOptions && chimMajEl.selectedOptions[0] ? chimMajEl.selectedOptions[0].text : '';
							const minText = chimMinEl.selectedOptions && chimMinEl.selectedOptions[0] ? chimMinEl.selectedOptions[0].text : '';
							if (majText || minText) {
								const cleanMaj = majText ? majText.replace(/^\s*\d+\s*[—-]?\s*/,'').trim() : '';
								const cleanMin = minText ? minText.replace(/^\s*\d+\s*[\-]?\s*/,'').trim() : '';
								app.chimney_full_words = [cleanMaj, cleanMin].filter(Boolean).join(' / ');
							}
						} catch (e) {}
					}
				}
			}
			
			// Get data from second materials table (own/shared, condition, shielding, label, location, flue_pipe_type)
			if (materialsTables.length >= 2) {
				const secondTable = materialsTables[1].querySelector('tbody tr');
				if (secondTable) {
					const ownSharedEl = secondTable.querySelector('select[name="own_shared"]');
					if (ownSharedEl) app.own_shared = ownSharedEl.value || '';
					const chimCondEl = secondTable.querySelector('select[name="chimney_condition"]');
					if (chimCondEl) app.chimney_condition = chimCondEl.value || '';
					const shieldingEl = secondTable.querySelector('select[name="shielding"]');
					if (shieldingEl) app.shielding = shieldingEl.value || '';
					const labelEl = secondTable.querySelector('select[name="label"]');
					if (labelEl) app.label = labelEl.value || '';
					const locationEl = secondTable.querySelector('select[name="location"]');
					if (locationEl) app.location = locationEl.value || '';
					const fluePipeTypeEl = secondTable.querySelector('select[name="flue_pipe_type"]');
					if (fluePipeTypeEl) app.flue_pipe_type = fluePipeTypeEl.value || '';
				}
			}
			
			// Collect clearances for this section
			const clearances = collectClearancesData(clearancesTable);
			
			// Collect remarks for this section
			const remarks = remarksField ? remarksField.value || '' : '';
			
			// Each section produces one entry with one appliance
			entries.push({ appliances: [app], clearances, remarks });
		});
		
		return entries;
	}
	
	if (shareBtn) {
		shareBtn.addEventListener('click', async function () {
			try {
				// Meta info
				const metaRaw = collectFormData();
				const meta = { policy: metaRaw.policy || '', survey_date: metaRaw.survey_date || '', completed_by: metaRaw.completed_by || '' };
				// Collect all appliance sections
				const entries = collectAllSections();

				if (!entries.length) {
					alert('No appliance entries to save.');
					return;
				}

				const { blob, suggestedName } = await generatePdfBlobForAll(meta, entries);
				const file = new File([blob], suggestedName, { type: 'application/pdf' });

				// Simple mobile detection
				const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
				const canShareFiles = typeof navigator !== 'undefined' && navigator.canShare && navigator.canShare({ files: [file] });

				if (isMobile && canShareFiles) {
					await navigator.share({ files: [file], title: suggestedName, text: 'Wood App Form' });
				} else {
					await saveBlobWithPicker(blob, suggestedName);
					alert('PDF generated as "' + suggestedName + '". If no dialog appeared, check your Downloads folder.');
				}
			} catch (err) {
				console.error('Share/Save failed', err);
				alert('Share/Save failed: ' + (err && err.message ? err.message : err));
			}
		});
	}

		// Preview button: generate combined PDF and show in modal iframe
		const previewBtn = document.getElementById('previewBtn');
		const previewModal = document.getElementById('previewModal');
		const previewFrame = document.getElementById('previewFrame');
		const closePreviewBtn = document.getElementById('closePreviewBtn');
		const downloadPreviewBtn = document.getElementById('downloadPreviewBtn');

		if (previewBtn && previewModal && previewFrame) {
			previewBtn.addEventListener('click', async function () {
				try {
					const metaRaw = collectFormData();
					const meta = { policy: metaRaw.policy || '', survey_date: metaRaw.survey_date || '', completed_by: metaRaw.completed_by || '' };
					// Collect all appliance sections
					const entries = collectAllSections();
					
					if (!entries.length) { 
						alert('No appliance entries to preview.'); 
						return; 
					}
					
					const { blob, suggestedName } = await generatePdfBlobForAll(meta, entries);
					
					// Detect mobile device - if mobile, open PDF in new tab for preview
					const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
					
					if (isMobile) {
						// On mobile, open PDF in a new tab/window for preview
						const url = URL.createObjectURL(blob);
						const newWindow = window.open(url, '_blank');
						if (!newWindow) {
							// If popup blocked, fall back to download
							const a = document.createElement('a');
							a.href = url;
							a.download = suggestedName;
							document.body.appendChild(a);
							a.click();
							a.remove();
						}
						// Clean up URL after a delay
						setTimeout(() => URL.revokeObjectURL(url), 60000);
					} else {
						// On desktop, show preview modal
						const url = URL.createObjectURL(blob);
						previewFrame.src = url;
						previewModal.setAttribute('aria-hidden', 'false');
						// store URL for later revoke
						previewModal._previewUrl = url;
						previewModal._previewSuggestedName = suggestedName;
						if (downloadPreviewBtn) {
							downloadPreviewBtn.onclick = function () {
								const a = document.createElement('a');
								a.href = url;
								a.download = suggestedName;
								document.body.appendChild(a);
								a.click();
								a.remove();
							};
						}
					}
				} catch (err) {
					console.error('Preview generation failed', err);
					alert('Preview generation failed: ' + (err && err.message ? err.message : err));
				}
			});

			// Close handler
			if (closePreviewBtn) closePreviewBtn.addEventListener('click', function () {
				if (previewFrame) previewFrame.src = '';
				previewModal.setAttribute('aria-hidden', 'true');
				if (previewModal._previewUrl) { URL.revokeObjectURL(previewModal._previewUrl); previewModal._previewUrl = null; }
			});

			// close when clicking outside modal content
			previewModal.addEventListener('click', function (e) {
				if (e.target === previewModal) {
					if (closePreviewBtn) closePreviewBtn.click();
				}
			});
		}

	async function createAndSavePdf(data) {
			// Use jsPDF (UMD exposes window.jspdf.jsPDF)
			const { jsPDF } = window.jspdf || {};
			if (!jsPDF) throw new Error('jsPDF library not loaded');

			const doc = new jsPDF({ unit: 'pt', format: 'a4' });
			const margin = 40;
			const pageWidth = doc.internal.pageSize.getWidth();
			let cursorY = margin;

			// Soft header band and Title (subtle coloring)
			doc.setFillColor(246, 250, 255);
			doc.rect(0, 0, pageWidth, 64, 'F');
			doc.setFontSize(18);
			doc.setFont('helvetica', 'bold');
			doc.setTextColor(22, 55, 92);
			doc.text('WOOD APP FORM', pageWidth / 2, cursorY, { align: 'center' });
			doc.setTextColor(0, 0, 0);
			cursorY += 26;

			// Small header table (policy, date, completed by)
			const policy = data.policy || '';
			const surveyDate = data.survey_date || '';
			const completedBy = data.completed_by || '';

			doc.setFontSize(10);
			doc.setFont('helvetica', 'normal');
			doc.autoTable({
				startY: cursorY,
				head: [['Policy # or Name', 'Survey date', 'Completed by']],
				body: [[policy, surveyDate, completedBy]],
				theme: 'grid',
				styles: { fontSize: 10 },
				headStyles: { fillColor: [225, 235, 245], textColor: 22, halign: 'center' },
				columnStyles: { 0: { cellWidth: 120 }, 1: { cellWidth: 120 }, 2: { cellWidth: 120 } }
			});
			cursorY = doc.lastAutoTable.finalY + 12;

			// Appliances table — use consistent columns
			// chimneyLegend collected here for use later in Notes (declare in outer scope)
			let chimneyLegend = [];
			if (Array.isArray(data.appliances) && data.appliances.length) {
				const appHead = ['Type', 'Make', 'Model', 'Installed By', 'Chimney Code', 'Own/Shared', 'Chimney Condition', 'Shielding', 'Label', 'Location', 'Flue Pipe Type'];
				// Collect chimney legend lines while building appliance rows so we can add the full wording to the Notes
				const appBody = data.appliances.map((app, idx) => {
					// compute chimney code from possible fields collected in the row
					const maj = app.chimney_major || app['chimney_major'] || '';
					const min = app.chimney_minor || app['chimney_minor'] || '';
					let chimneyCode = app.chimney_code || app['chimney_code'] || '';
					if ((!chimneyCode || chimneyCode === '') && maj) {
						chimneyCode = min ? `${maj}.${min}` : `${maj}`;
					}

					// Try to read the full option text from the DOM row if available (preserve human-friendly wording)
					let chimneyFullWords = '';
					try {
						if (materialsTable) {
							const rows = Array.from(materialsTable.querySelectorAll('tr'));
							const row = rows[idx];
							if (row) {
								const majSel = row.querySelector('select[name="chimney_major"]');
								const minSel = row.querySelector('select[name="chimney_minor"]');
								const majText = majSel && majSel.selectedOptions && majSel.selectedOptions[0] ? majSel.selectedOptions[0].text : '';
								const minText = minSel && minSel.selectedOptions && minSel.selectedOptions[0] ? minSel.selectedOptions[0].text : '';
								if (majText || minText) {
									const cleanMaj = majText ? majText.replace(/^\s*\d+\s*[—-]?\s*/,'').trim() : '';
									const cleanMin = minText ? minText.replace(/^\s*\d+\s*[\-]?\s*/,'').trim() : '';
									chimneyFullWords = [cleanMaj, cleanMin].filter(Boolean).join(' / ');
								}
							}
						}
					} catch (e) {
						// ignore DOM lookup errors
					}

					if (chimneyCode) {
						chimneyLegend.push({ code: chimneyCode, words: chimneyFullWords });
					}

						return [
							app.type || app['type'] || (app['col0'] || ''),
							app.make || '',
							app.model || '',
							app.installed_by || app['installed_by'] || '',
							chimneyCode,
							app.own_shared || app['own_shared'] || '',
							app.chimney_condition || app['chimney_condition'] || '',
							(app.shielding === true || app.shielding === 'yes' || app.shielding === 'Yes') ? 'Yes' : (app.shielding === 'no' || app.shielding === false ? 'No' : (app.shielding || '')),
							app.label || '',
							app.location || ''
						];
				});

				doc.setFontSize(12);
				doc.setFont('helvetica', 'bold');
				doc.text('Appliance Details', margin, cursorY);
				cursorY += 8;

					doc.autoTable({
						startY: cursorY,
						head: [appHead],
						body: appBody,
						styles: { fontSize: 9 },
						headStyles: { fillColor: [235, 245, 255], textColor: 22 },
						theme: 'striped',
						columnStyles: { 0: { cellWidth: 60 }, 1: { cellWidth: 70 }, 2: { cellWidth: 70 }, 3: { cellWidth: 70 } }
					});
				cursorY = doc.lastAutoTable.finalY + 12;
			}

			// Clearances table — read directly from DOM to preserve order and special rows
			if (clearancesTable) {
				const hasShielded = Array.from(clearancesTable.querySelectorAll('thead th')).some(th => th.textContent.trim() === 'Shielded');
				const head = hasShielded ? ['Clearances from', 'Required', 'Actual', 'Shielded'] : ['Clearances from', 'Required', 'Actual'];

			const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
			const body = rows.map(r => {
				const first = r.querySelector('td');
				if (!first) return null;
				const colspan = first.getAttribute('colspan');
				const label = first.textContent.trim();
				
				// Capture row styling class for PDF coloring
				let rowClass = '';
				if (r.classList.contains('row-positive')) rowClass = 'positive';
				else if (r.classList.contains('row-negative')) rowClass = 'negative';
				else if (r.classList.contains('row-caution')) rowClass = 'caution';
				
				if (colspan && parseInt(colspan) > 1) {
					return { label: label, required: '', actual: '', shielded: '', _isHeader: true, _rowClass: rowClass };
				}
				const cells = r.querySelectorAll('td');
				const reqInput = cells[1] ? cells[1].querySelector('input, select, textarea') : null;
				const actInput = cells[2] ? cells[2].querySelector('input, select, textarea') : null;
				const req = reqInput ? (reqInput.value || '') : (cells[1] ? cells[1].textContent.trim() : '');
				const act = actInput ? (actInput.value || '') : (cells[2] ? cells[2].textContent.trim() : '');
				let shieldVal = '';
				if (hasShielded) {
					const shieldCell = cells[3];
					if (shieldCell) {
						const chk = shieldCell.querySelector('input[type="checkbox"]');
						shieldVal = chk ? (chk.checked ? 'Yes' : 'No') : shieldCell.textContent.trim();
					}
				}
				return { label, required: req, actual: act, shielded: shieldVal, _isHeader: false, _rowClass: rowClass };
			}).filter(Boolean);				if (body.length) {
					doc.setFontSize(12);
					doc.setFont('helvetica', 'bold');
					doc.text('Measurements & Clearances', margin, cursorY);
					cursorY += 8;

					// convert to array rows for autoTable, and use didParseCell to handle header rows
					const atBody = body.map(b => hasShielded ? [b.label, b.required, b.actual, b.shielded] : [b.label, b.required, b.actual]);

					doc.autoTable({
						startY: cursorY,
						head: [head],
						body: atBody,
						styles: { fontSize: 9 },
						headStyles: { fillColor: [235, 245, 255], textColor: 22 },
						theme: 'grid',
						didParseCell: function (dataCell) {
							// dataCell.row.raw gives the raw array element; detect header-style rows where required+actual are empty
							const raw = dataCell.row && dataCell.row.raw;
							if (!raw) return;
							// raw here is an array like [label, req, act, (shield)]
							const isHeader = raw[1] === '' && raw[2] === '' && (hasShielded ? raw[3] === '' : true);
							if (isHeader) {
								// make the first column span all
								if (dataCell.column.index === 0) {
									dataCell.cell.colSpan = hasShielded ? 4 : 3;
									// Only apply blue background to Floor Pad row
									const label = raw[0] || '';
									if (label.toLowerCase().includes('floor pad')) {
										dataCell.cell.styles.fillColor = [235, 245, 255];
										dataCell.cell.styles.textColor = 22;
									}
									dataCell.cell.styles.halign = 'left';
					} else {
						dataCell.cell.text = '';
					}
				} else {
					// Apply row coloring based on _rowClass (only to body rows, not header)
					if (dataCell.row.section === 'body') {
						const bodyItem = body[dataCell.row.index];
						if (bodyItem && bodyItem._rowClass) {
							if (bodyItem._rowClass === 'positive') {
								dataCell.cell.styles.fillColor = [230, 255, 230];
							} else if (bodyItem._rowClass === 'negative') {
								dataCell.cell.styles.fillColor = [255, 235, 235];
							} else if (bodyItem._rowClass === 'caution') {
								dataCell.cell.styles.fillColor = [255, 250, 230];
							}
						}
					}
				}
			}
		});					cursorY = doc.lastAutoTable.finalY + 12;
				}
			}

			// Notes / Remarks (include chimney code wording legend)
			doc.setFontSize(12);
			doc.setFont('helvetica', 'bold');
			doc.text('Notes', margin, cursorY);
			cursorY += 12;
			doc.setFont('helvetica', 'normal');
			doc.setFontSize(10);
			const notesText = (data.remarks || '').trim();
			const notesLines = notesText ? doc.splitTextToSize(notesText, pageWidth - margin * 2) : [];
			if (notesLines.length) {
				doc.text(notesLines, margin, cursorY);
				cursorY += notesLines.length * 12 + 8;
			}

			// Chimney code legend: dedupe by code and show the selected wording
			if (Array.isArray(chimneyLegend) && chimneyLegend.length) {
				// build unique map
				const uniq = {};
				chimneyLegend.forEach(item => {
					if (!item || !item.code) return;
					if (!uniq[item.code]) uniq[item.code] = item.words || '';
					else if (!uniq[item.code] && item.words) uniq[item.code] = item.words;
				});

				const legendLines = [];
				for (const code of Object.keys(uniq)) {
					const words = uniq[code];
					if (words) legendLines.push(`${code} — ${words}`);
					else legendLines.push(`${code}`);
				}

				if (legendLines.length) {
					doc.setFont('helvetica', 'bold');
					doc.text('Chimney Code Legend', margin, cursorY);
					cursorY += 12;
					doc.setFont('helvetica', 'normal');
					const wrapped = doc.splitTextToSize(legendLines.join('\n'), pageWidth - margin * 2);
					doc.text(wrapped, margin, cursorY);
					cursorY += wrapped.length * 12 + 8;
				}
			}

			// (Signatures removed as requested)

			// finalize PDF
			const blob = doc.output('blob');

			// suggested filename: policy # - date - Wood app form.pdf
			const safePolicy = sanitizeFilename(policy || 'policy');
			const safeDate = sanitizeFilename(surveyDate || (new Date()).toISOString().slice(0,10));
			const suggestedName = `${safePolicy} - ${safeDate} - Wood app form.pdf`;

			await saveBlobWithPicker(blob, suggestedName);
		}

	if (form) {
		form.addEventListener('submit', async function (ev) {
			ev.preventDefault();
			const data = collectFormData();
			console.log('Wood App Form data:', data);

			// create a PDF from the collected data
			try {
				await createAndSavePdf(data);
			} catch (err) {
				console.error('PDF generation / save failed', err);
				alert('Unable to save PDF: ' + (err && err.message ? err.message : err));
			}
		});
	}

	// Register service worker for PWA (if supported)
	if ('serviceWorker' in navigator) {
		window.addEventListener('load', function () {
			navigator.serviceWorker.register('sw.js').then(reg => {
				console.log('ServiceWorker registered:', reg.scope);
			}).catch(err => {
				console.warn('ServiceWorker registration failed:', err);
			});
		});
	}

	// chimney info button toggle
	const infoBtn = document.getElementById('chimneyInfoBtn');
	const infoBox = document.getElementById('chimneyInfo');
	if (infoBtn && infoBox) {
		infoBtn.addEventListener('click', function () {
			const shown = infoBox.style.display !== 'none';
			infoBox.style.display = shown ? 'none' : 'block';
			infoBtn.setAttribute('aria-expanded', String(!shown));
		});
	}

	// Shielding behaviour: toggle 'Shielded' column in clearances table
	const shieldingSelect = document.getElementById('shielding');

	function makeSafeName(name) {
		return name.replace(/[^a-z0-9_]/gi, '_');
	}

	function addShieldedColumn() {
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		// if Shielded header already exists, do nothing
		if (Array.from(theadRow.children).some(th => th.textContent.trim() === 'Shielded')) return;

		// insert header at position 3 (after Actual) to keep a consistent column order
		const th = document.createElement('th');
		th.textContent = 'Shielded';
		const insertIndex = 3; // 0-based
		if (theadRow.children.length > insertIndex) theadRow.insertBefore(th, theadRow.children[insertIndex]);
		else theadRow.appendChild(th);

		// add checkbox cell for each tbody row (adjust colspan rows by increasing colspan)
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			const colspan = firstCell.getAttribute('colspan') ? parseInt(firstCell.getAttribute('colspan')) : 1;
			if (colspan > 1) {
				// increase colspan to account for new column
				firstCell.setAttribute('colspan', colspan + 1);
				return;
			}

			// derive a name for the shielded field from existing inputs in the row
			const ctrl = row.querySelector('input[name], select[name], textarea[name]');
			let baseName = 'row';
			if (ctrl && ctrl.name) {
				baseName = ctrl.name.replace(/(_required|_actual)$/,'');
			} else {
				baseName = makeSafeName(firstCell.textContent.trim().toLowerCase());
			}
			const shieldName = `shielded_${baseName}`;

		const td = document.createElement('td');
		const input = document.createElement('input');
		input.type = 'checkbox';
		input.name = shieldName;
		input.className = 'shielded-checkbox';
		
		// Add event listener to update required value for flue pipe rows
		input.addEventListener('change', function() {
			updateFluePipeRequiredValue(row, input);
		});
		
		td.appendChild(input);
		// insert at the correct index in the row
		if (row.children.length > insertIndex) row.insertBefore(td, row.children[insertIndex]);
		else row.appendChild(td);
	});
}	function removeShieldedColumn() {
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		// find Shielded header index
		const ths = Array.from(theadRow.children);
		let shieldIndex = -1;
		for (let i = 0; i < ths.length; i++) {
			if (ths[i].textContent.trim() === 'Shielded') { shieldIndex = i; break; }
		}
		if (shieldIndex >= 0) theadRow.removeChild(ths[shieldIndex]);

		// remove or reduce cells from each tbody row
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			const colspan = firstCell.getAttribute('colspan') ? parseInt(firstCell.getAttribute('colspan')) : 1;
			if (colspan > 1) {
				// reduce colspan
				const newCol = Math.max(1, colspan - 1);
				firstCell.setAttribute('colspan', newCol);
				return;
			}
			// otherwise remove the td at shieldIndex if present
			const cells = row.querySelectorAll('td');
			if (shieldIndex >= 0 && cells.length > shieldIndex) {
				const cell = cells[shieldIndex];
				if (cell) {
					// remove if it's the shielded checkbox or empty
					if (cell.querySelector && (cell.querySelector('.shielded-checkbox') || cell.innerHTML.trim() === '')) {
						cell.remove();
					}
				}
			}
		});
	}

	// watch the shielding select change (it exists in the appliances table row)
	if (shieldingSelect) {
		shieldingSelect.addEventListener('change', function (e) {
			const val = (shieldingSelect.value || '').toString().toLowerCase();
			if (val === 'yes') addShieldedColumn();
			else removeShieldedColumn();
		});
		// initialize on load (case-insensitive)
		if ((shieldingSelect.value || '').toString().toLowerCase() === 'yes') addShieldedColumn();
	}

	// Apply visibility rules based on appliance type
	const typeSelect = document.getElementById('type');

	function findFirstCellText(row) {
		const td = row.querySelector('td');
		return td ? td.textContent.trim().toLowerCase() : '';
	}

	function getAllClearanceRows() {
		if (!clearancesTable) return [];
		return Array.from(clearancesTable.querySelectorAll('tbody tr'));
	}

	function hideRowsByKeywords(keywords) {
		const rows = getAllClearanceRows();
		rows.forEach(r => {
			const text = findFirstCellText(r);
			const match = keywords.some(k => text.includes(k));
			if (match) {
				r.style.display = 'none';
				Array.from(r.querySelectorAll('input, select, textarea')).forEach(el => {
					if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
					else el.value = '';
				});
			} else {
				r.style.display = '';
			}
		});
	}

	function removeFacingRows() {
		if (!clearancesTable) return;
		const rows = getAllClearanceRows();
		rows.forEach(r => {
			const text = findFirstCellText(r);
			if (text === 'left facing' || text === 'right facing') r.remove();
		});
	}

	function facingRowExists(label) {
		if (!clearancesTable) return false;
		return getAllClearanceRows().some(r => findFirstCellText(r) === label.toLowerCase());
	}

	function addFacingRowsAfterRightSide() {
		if (!clearancesTable) return;
		// don't duplicate
		if (facingRowExists('left facing') || facingRowExists('right facing')) return;

		const rows = getAllClearanceRows();
		let insertAfter = null;
		for (const r of rows) {
			const txt = findFirstCellText(r);
			if (txt === 'right side') { insertAfter = r; break; }
		}
		// if not found, append to tbody end
		const tbody = clearancesTable.querySelector('tbody');
		const createFacingRow = (label) => {
			const tr = document.createElement('tr');
			const tdLabel = document.createElement('td');
			tdLabel.textContent = label;
			const tdReq = document.createElement('td');
			const reqInput = document.createElement('input');
			reqInput.type = 'text';
			reqInput.name = `${label.toLowerCase().replace(/\s+/g,'_')}_required`;
			tdReq.appendChild(reqInput);
			const tdAct = document.createElement('td');
			const actInput = document.createElement('input');
			actInput.type = 'text';
			actInput.name = `${label.toLowerCase().replace(/\s+/g,'_')}_actual`;
			tdAct.appendChild(actInput);
			tr.appendChild(tdLabel);
			tr.appendChild(tdReq);
			tr.appendChild(tdAct);

			// if shielded column present, append shield checkbox cell
			const theadRow = clearancesTable.querySelector('thead tr');
			if (Array.from(theadRow.children).some(th => th.textContent.trim() === 'Shielded')) {
				const tdShield = document.createElement('td');
				const chk = document.createElement('input');
				chk.type = 'checkbox';
				chk.name = `shielded_${label.toLowerCase().replace(/\s+/g,'_')}`;
				chk.className = 'shielded-checkbox';
				tdShield.appendChild(chk);
				tr.appendChild(tdShield);
			}

			return tr;
		};

		const leftTr = createFacingRow('Left facing');
		const rightTr = createFacingRow('Right facing');
		if (insertAfter && insertAfter.parentNode) {
			insertAfter.parentNode.insertBefore(leftTr, insertAfter.nextSibling);
			insertAfter.parentNode.insertBefore(rightTr, leftTr.nextSibling);
		} else {
			tbody.appendChild(leftTr);
			tbody.appendChild(rightTr);
		}
	}

function applyTypeRules(val) {
	if (!clearancesTable) return;
	// normalize
	const v = (val || '').toString().toLowerCase();

	// start by showing all rows and restore their default values
	getAllClearanceRows().forEach(r => {
		r.style.display = '';
		Array.from(r.querySelectorAll('input, select, textarea')).forEach(el => {
			if (el.type === 'checkbox' || el.type === 'radio') {
				el.checked = false;
			} else if (el.defaultValue !== undefined && el.defaultValue !== '') {
				el.value = el.defaultValue;
			}
		});
	});

	// remove any previously added facing rows to avoid duplicates; we'll add if needed
	removeFacingRows();

	// ensure labels are in their default 'Flue pipe ...' form before applying rules
	renameFlueToChimney(false);		// rename flue -> chimney for outdoor boiler
		function renameFlueToChimney(shouldRename) {
			const mapping = [
				['flue pipe back', 'Chimney back'],
				['flue pipe side', 'Chimney side'],
				['flue pipe ceiling', 'Chimney ceiling']
			];
			const rows = getAllClearanceRows();
			rows.forEach(r => {
				const first = r.querySelector('td');
				if (!first) return;
				const text = first.textContent.trim().toLowerCase();
				mapping.forEach(([from, to]) => {
					if (shouldRename) {
						if (text.includes(from)) {
							first.textContent = to;
						}
					} else {
						// revert if currently renamed
						if (text.includes(to.toLowerCase())) {
							// restore original casing
							first.textContent = from.charAt(0).toUpperCase() + from.slice(1);
						}
					}
				});
			});
		}

		// rules per type
		switch (v) {
			case 'kitchen wood range':
				hideRowsByKeywords(['plenum','mantel']);
				break;
			case 'insert':
				hideRowsByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum']);
				addFacingRowsAfterRightSide();
				break;
			case 'furnace':
				hideRowsByKeywords(['left corner','right corner','mantel','top']);
				break;
			case 'boiler':
				hideRowsByKeywords(['plenum','mantel']);
				break;
			case 'factory built fireplace':
				hideRowsByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum','floor pad rear']);
				addFacingRowsAfterRightSide();
				break;
			case 'pellet stove/insert':
				hideRowsByKeywords(['plenum']);
				break;
			case 'hearth':
				hideRowsByKeywords(['plenum']);
				break;
			case 'outdoor wood boiler':
				// rename flue labels to chimney and apply hides
				renameFlueToChimney(true);
				hideRowsByKeywords(['plenum','mantel']);
				break;
			case 'masonry fireplace':
				hideRowsByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum','floor pad rear']);
				addFacingRowsAfterRightSide();
				break;
			case 'stove':
				hideRowsByKeywords(['plenum','mantel']);
				break;
			default:
				// no special hiding
				break;
		}
	}

	if (typeSelect) {
		typeSelect.addEventListener('change', function () { applyTypeRules(typeSelect.value); });
		// initialize on load
		applyTypeRules(typeSelect.value);
	}

	// Add listener for flue pipe type changes in first appliance
	const fluePipeTypeSelect = document.getElementById('flue_pipe_type');
	if (fluePipeTypeSelect) {
		fluePipeTypeSelect.addEventListener('change', function () {
			const section = fluePipeTypeSelect.closest('.appliance-section');
			if (section) {
				applyFluePipeTypeClearances(section, fluePipeTypeSelect.value);
			}
		});
	}

	if (resetBtn) {
		resetBtn.addEventListener('click', function () {
			const modal = document.getElementById('resetModal');
			if (modal) {
				modal.setAttribute('aria-hidden', 'false');
			} else {
				// fallback to immediate reset if modal missing
				doReset();
			}
		});

		// modal buttons
		const resetModal = document.getElementById('resetModal');
		const confirmReset = document.getElementById('confirmReset');
		const cancelReset = document.getElementById('cancelReset');
		function closeModal() { if (resetModal) resetModal.setAttribute('aria-hidden', 'true'); }
		if (cancelReset) cancelReset.addEventListener('click', function () { closeModal(); });
		if (confirmReset) confirmReset.addEventListener('click', function () { closeModal(); doReset(); });
	}

	function doReset() {
		if (form) form.reset();
		// reset survey date to today
		if (surveyInput) surveyInput.value = new Date().toISOString().slice(0,10);
		
		// Remove shielded column from first appliance if present
		removeShieldedColumn();
		
		// Remove all appliance sections except the first one
		const container = document.getElementById('appliancesContainer');
		if (container) {
			const sections = Array.from(container.querySelectorAll('.appliance-section'));
			sections.forEach((section, index) => {
				if (index === 0) {
					// Clear first section
					section.querySelectorAll('input[type="text"], input[type="number"], input[type="email"], input[type="date"], textarea').forEach(inp => {
						if (inp.hasAttribute('value') && inp.getAttribute('value')) {
							inp.value = inp.getAttribute('value');
						} else {
							inp.value = '';
						}
					});
					section.querySelectorAll('select').forEach(sel => {
						const defaultSelected = sel.querySelector('option[selected]');
						if (defaultSelected) {
							sel.value = defaultSelected.value;
						} else {
							sel.selectedIndex = 0;
						}
					});
					section.querySelectorAll('input[type="checkbox"], input[type="radio"]').forEach(inp => {
						inp.checked = false;
					});
					section.querySelectorAll('tbody tr').forEach(r => {
						r.classList.remove('row-positive','row-negative','row-caution');
						r.style.display = '';
					});
				} else {
					// Remove extra sections
					section.remove();
				}
			});
		}
	}
});








