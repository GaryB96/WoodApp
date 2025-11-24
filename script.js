document.addEventListener('DOMContentLoaded', function () {
	const form = document.getElementById('woodAppForm');
	let result = document.getElementById('result');
	const addMaterialBtn = document.getElementById('addMaterial');
	const materialsTableElem = document.getElementById('materialsTable');
	const materialsTable = materialsTableElem ? materialsTableElem.querySelector('tbody') : null;
	const resetBtn = document.getElementById('resetBtn');

	// we will not create or use a visible result area — PDF save will be offered on submit

	// Only wire add/remove if the buttons/table exist
	if (addMaterialBtn && materialsTable) {
		addMaterialBtn.addEventListener('click', function () {
			const tr = document.createElement('tr');
			tr.innerHTML = `
				<td data-label="Type"><input type="text" name="type"></td>
				<td data-label="Make"><input type="text" name="make"></td>
				<td data-label="Model"><input type="text" name="model"></td>
				<td data-label="Installed By"><input type="text" name="installed_by"></td>
				<td data-label="Chimney Code"><input type="text" name="chimney_code"></td>
				<td data-label="Own/Shared"><input type="text" name="own_shared"></td>
				<td data-label="Chimney Condition"><input type="text" name="chimney_condition"></td>
				<td data-label="Actions"><button type="button" class="remove-row">×</button></td>
			`;
			materialsTable.appendChild(tr);
		});

		materialsTable.addEventListener('click', function (e) {
			if (e.target && e.target.classList.contains('remove-row')) {
				const row = e.target.closest('tr');
				if (row) row.remove();
			}
		});
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
				head: [['Policy #', 'Survey date', 'Completed by']],
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
				const appHead = ['Type', 'Make', 'Model', 'Installed By', 'Chimney Code', 'Own/Shared', 'Chimney Condition', 'Shielding', 'Label'];
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
						app.label || ''
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
					if (colspan && parseInt(colspan) > 1) {
						return { label: label, required: '', actual: '', shielded: '', _isHeader: true };
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
					return { label, required: req, actual: act, shielded: shieldVal, _isHeader: false };
				}).filter(Boolean);

				if (body.length) {
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
						headStyles: { fillColor: [240, 240, 240], textColor: 22 },
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
									dataCell.cell.styles.fillColor = [235, 245, 255];
									dataCell.cell.styles.textColor = 22;
									dataCell.cell.styles.halign = 'left';
								} else {
									dataCell.cell.text = '';
								}
							}
						}
					});

					cursorY = doc.lastAutoTable.finalY + 12;
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
	const clearancesTable = document.getElementById('clearancesTable');

	function makeSafeName(name) {
		return name.replace(/[^a-z0-9_]/gi, '_');
	}

	function addShieldedColumn() {
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		// if Shielded header already exists, do nothing
		if (Array.from(theadRow.children).some(th => th.textContent.trim() === 'Shielded')) return;

		// add header
		const th = document.createElement('th');
		th.textContent = 'Shielded';
		theadRow.appendChild(th);

		// add checkbox cell for each tbody row (skip rows that span all columns)
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const firstCell = row.querySelector('td');
			if (!firstCell) return;
			if (firstCell.getAttribute('colspan') && parseInt(firstCell.getAttribute('colspan')) > 1) {
				// add an empty cell to preserve structure (or skip)
				const empty = document.createElement('td');
				empty.innerHTML = '';
				row.appendChild(empty);
				return;
			}

			// derive a name for the shielded field from existing inputs in the row
			const ctrl = row.querySelector('input[name], select[name], textarea[name]');
			let baseName = 'row';
			if (ctrl && ctrl.name) {
				baseName = ctrl.name.replace(/(_required|_actual)$/,'');
			} else {
				// fallback to using text content
				baseName = makeSafeName(firstCell.textContent.trim().toLowerCase());
			}
			const shieldName = `shielded_${baseName}`;

			const td = document.createElement('td');
			const input = document.createElement('input');
			input.type = 'checkbox';
			input.name = shieldName;
			input.className = 'shielded-checkbox';
			td.appendChild(input);
			row.appendChild(td);
		});
	}

	function removeShieldedColumn() {
		if (!clearancesTable) return;
		const theadRow = clearancesTable.querySelector('thead tr');
		// remove Shielded header if present
		const ths = Array.from(theadRow.children);
		for (let i = ths.length - 1; i >= 0; i--) {
			if (ths[i].textContent.trim() === 'Shielded') {
				theadRow.removeChild(ths[i]);
				break;
			}
		}

		// remove last cell from each row if it's a shielded checkbox (or remove empty placeholders)
		const rows = Array.from(clearancesTable.querySelectorAll('tbody tr'));
		rows.forEach(row => {
			const cells = row.querySelectorAll('td');
			if (!cells.length) return;
			const last = cells[cells.length - 1];
			if (!last) return;
			// if last cell contains a shielded-checkbox or is empty placeholder, remove it
			if (last.querySelector && last.querySelector('.shielded-checkbox')) {
				last.remove();
			} else if (last.innerHTML.trim() === '') {
				last.remove();
			}
		});
	}

	// watch the shielding select change (it exists in the appliances table row)
	if (shieldingSelect) {
		shieldingSelect.addEventListener('change', function (e) {
			const val = shieldingSelect.value;
			if (val === 'yes') addShieldedColumn();
			else removeShieldedColumn();
		});
		// initialize on load
		if (shieldingSelect.value === 'yes') addShieldedColumn();
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

		// start by showing all rows
		getAllClearanceRows().forEach(r => r.style.display = '');

		// remove any previously added facing rows to avoid duplicates; we'll add if needed
		removeFacingRows();

		// ensure labels are in their default 'Flue pipe ...' form before applying rules
		renameFlueToChimney(false);

		// rename flue -> chimney for outdoor boiler
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
			case 'range': // Kitchen wood range
				hideRowsByKeywords(['plenum','mantel']);
				break;
			case 'insert':
				hideRowsByKeywords(['rear','flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum']);
				addFacingRowsAfterRightSide();
				break;
			case 'furnace':
				hideRowsByKeywords(['mantel']);
				break;
			case 'boiler':
				hideRowsByKeywords(['plenum','mantel']);
				break;
			case 'factorybuilt':
				hideRowsByKeywords(['flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum']);
				addFacingRowsAfterRightSide();
				break;
			case 'pellet':
				hideRowsByKeywords(['plenum']);
				break;
			case 'hearth':
				hideRowsByKeywords(['plenum']);
				break;
			case 'outdoorboiler':
				// rename flue labels to chimney and apply hides
				renameFlueToChimney(true);
				hideRowsByKeywords(['left corner','right corner','plenum','mantel']);
				break;
			case 'fireplace': // masonry fireplace
				hideRowsByKeywords(['flue pipe back','flue pipe side','flue pipe ceiling','left corner','right corner','plenum']);
				addFacingRowsAfterRightSide();
				break;
			case 'stove':
				// keep previous behavior: hide plenum
				hideRowsByKeywords(['plenum']);
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

	if (resetBtn) {
		resetBtn.addEventListener('click', function () {
			if (form) form.reset();
			// clear dynamic rows if table exists
			if (materialsTable) {
				const rows = Array.from(materialsTable.querySelectorAll('tr'));
				rows.forEach((r, i) => {
					// clear inputs in first row, remove others
					if (i === 0) {
						Array.from(r.querySelectorAll('input, select, textarea')).forEach(inp => {
							if (inp.type === 'checkbox') inp.checked = false;
							else inp.value = '';
						});
					} else {
						r.remove();
					}
				});
			}
			// no result area to update; PDF/save flow handles output
		});
	}
});
