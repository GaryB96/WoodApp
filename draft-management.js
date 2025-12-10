// ============================================
// DRAFT MANAGEMENT SYSTEM
// ============================================
// This file should be included after script.js in index.html

document.addEventListener('DOMContentLoaded', function() {
	const form = document.getElementById('woodAppForm');
	const saveDraftBtn = document.getElementById('saveDraftBtn');
	const loadDraftBtn = document.getElementById('loadDraftBtn');
	const manageDraftsBtn = document.getElementById('manageDraftsBtn');
	const draftIndicator = document.getElementById('draftIndicator');
	const manageDraftsModal = document.getElementById('manageDraftsModal');
	const closeDraftsModal = document.getElementById('closeDraftsModal');
	const draftsList = document.getElementById('draftsList');
	const shareBtn = document.getElementById('shareBtn');

	const DRAFTS_KEY = 'woodapp_drafts';

	let currentlyLoadedDraftId = null; // Track which draft was loaded

	// Get all drafts from localStorage
	function getAllDrafts() {
		try {
			const stored = localStorage.getItem(DRAFTS_KEY);
			return stored ? JSON.parse(stored) : [];
		} catch (e) {
			console.error('Error loading drafts:', e);
			return [];
		}
	}

	// Save drafts array to localStorage
	function saveAllDrafts(drafts) {
		try {
			localStorage.setItem(DRAFTS_KEY, JSON.stringify(drafts));
			updateDraftUI();
		} catch (e) {
			console.error('Error saving drafts:', e);
			alert('Failed to save draft. Storage may be full.');
		}
	}

	// Collect complete form state including all appliances
	function collectFormState() {
		// Get all form inputs with their current values
		const formData = {};
		const formElements = form.querySelectorAll('input, select, textarea');
		formElements.forEach(elem => {
			// Create unique key that includes any parent section identifier
			let key = elem.name;
			if (!key) return;
			
			// For elements inside appliance sections, add section number to key
			const applianceSection = elem.closest('.appliance-section');
			if (applianceSection) {
				const sectionNum = applianceSection.getAttribute('data-appliance-num');
				if (sectionNum) {
					key = `section${sectionNum}_${elem.name}`;
				}
			}
			
			if (elem.type === 'checkbox') {
				formData[key] = elem.checked;
			} else if (elem.type === 'radio') {
				if (elem.checked) {
					formData[key] = elem.value;
				}
			} else {
				formData[key] = elem.value;
			}
		});

		// Store the number of appliance sections
		const container = document.getElementById('appliancesContainer');
		const sectionCount = container ? container.querySelectorAll('.appliance-section').length : 1;

		return {
			formData,
			sectionCount,
			timestamp: Date.now()
		};
	}

	// Restore form state
	function restoreFormState(state) {
		if (!state) return;

		// First, ensure we have the right number of appliance sections
		if (state.sectionCount) {
			const container = document.getElementById('appliancesContainer');
			if (container) {
				const currentSections = container.querySelectorAll('.appliance-section').length;
				const addBtn = document.getElementById('addAnotherBtn');
				
				// Add sections if needed
				if (addBtn && state.sectionCount > currentSections) {
					for (let i = currentSections; i < state.sectionCount; i++) {
						addBtn.click();
						// Small delay to allow DOM to update
					}
				}
				
				// Remove extra sections if needed
				if (state.sectionCount < currentSections) {
					const sections = container.querySelectorAll('.appliance-section');
					for (let i = state.sectionCount; i < currentSections; i++) {
						if (sections[i]) {
							sections[i].remove();
						}
					}
				}
			}
		}

		// Wait a moment for sections to be added/removed, then restore values
		setTimeout(() => {
			// Restore all form field values
			if (state.formData) {
				Object.keys(state.formData).forEach(key => {
					// Handle section-specific fields
					if (key.startsWith('section')) {
						const match = key.match(/^section(\d+)_(.+)$/);
						if (match) {
							const sectionNum = match[1];
							const fieldName = match[2];
							const section = document.querySelector(`.appliance-section[data-appliance-num="${sectionNum}"]`);
							if (section) {
								const elements = section.querySelectorAll(`[name="${fieldName}"]`);
								elements.forEach(elem => {
									if (elem.type === 'checkbox') {
										elem.checked = state.formData[key] === true;
									} else if (elem.type === 'radio') {
										elem.checked = elem.value === state.formData[key];
									} else {
										elem.value = state.formData[key] || '';
									}
									// Trigger change event to update dependent fields
									elem.dispatchEvent(new Event('change', { bubbles: true }));
								});
							}
						}
					} else {
						// Handle global form fields (policy, survey_date, completed_by)
						const elements = form.querySelectorAll(`[name="${key}"]`);
						elements.forEach(elem => {
							if (elem.type === 'checkbox') {
								elem.checked = state.formData[key] === true;
							} else if (elem.type === 'radio') {
								elem.checked = elem.value === state.formData[key];
							} else {
								elem.value = state.formData[key] || '';
							}
							// Trigger change event to update dependent fields
							elem.dispatchEvent(new Event('change', { bubbles: true }));
						});
					}
				});
			}
		}, 100);
	}

	// Save current form as draft
	function saveDraft() {
		const drafts = getAllDrafts();
		const state = collectFormState();
		
		// Prompt for draft name
		const policy = form.querySelector('[name="policy"]')?.value || '';
		const defaultName = policy ? `${policy} - ${new Date().toLocaleString()}` : `Draft - ${new Date().toLocaleString()}`;
		const name = prompt('Enter a name for this draft:', defaultName);
		
		if (name === null) return; // User cancelled
		
		const draft = {
			id: Date.now(),
			name: name || defaultName,
			timestamp: state.timestamp,
			state: state
		};
		
		drafts.push(draft);
		saveAllDrafts(drafts);
		alert('Draft saved successfully!');
	}

	// Load most recent draft
	function loadMostRecentDraft() {
		const drafts = getAllDrafts();
		if (drafts.length === 0) {
			alert('No drafts available.');
			return;
		}
		
		const mostRecent = drafts[drafts.length - 1];
		
		if (confirm(`Load draft "${mostRecent.name}"? This will replace current form data.`)) {
			restoreFormState(mostRecent.state);
		}
	}

	// Show manage drafts modal
	function showManageDrafts() {
		const drafts = getAllDrafts();
		if (drafts.length === 0) {
			alert('No saved drafts.');
			return;
		}
		
		draftsList.innerHTML = '';
		drafts.forEach((draft, index) => {
			const draftItem = document.createElement('div');
			draftItem.style.cssText = 'border:1px solid #ddd; padding:12px; margin-bottom:8px; border-radius:4px; background:#f9f9f9;';
			
			const timestamp = new Date(draft.timestamp).toLocaleString();
			draftItem.innerHTML = `
				<div style="display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap;">
					<div style="flex:1; min-width:200px;">
						<strong>${draft.name}</strong><br>
						<small style="color:#666;">${timestamp}</small>
					</div>
					<div style="display:flex; gap:4px;">
						<button class="load-draft-btn" data-index="${index}" style="padding:6px 12px; cursor:pointer;">Load</button>
						<button class="delete-draft-btn" data-index="${index}" style="padding:6px 12px; background:#d32f2f; color:white; cursor:pointer; border:none; border-radius:4px;">Delete</button>
					</div>
				</div>
			`;
			draftsList.appendChild(draftItem);
		});
		
		// Add event listeners
		draftsList.querySelectorAll('.load-draft-btn').forEach(btn => {
			btn.addEventListener('click', function() {
				const index = parseInt(this.getAttribute('data-index'));
				const draft = drafts[index];
				if (confirm(`Load draft "${draft.name}"? This will replace current form data.`)) {
					restoreFormState(draft.state);
					currentlyLoadedDraftId = draft.id; // Track which draft was loaded
					manageDraftsModal.setAttribute('aria-hidden', 'true');
				}
			});
		});
		
		draftsList.querySelectorAll('.delete-draft-btn').forEach(btn => {
			btn.addEventListener('click', function() {
				const index = parseInt(this.getAttribute('data-index'));
				const draft = drafts[index];
				if (confirm(`Delete draft "${draft.name}"?`)) {
					drafts.splice(index, 1);
					saveAllDrafts(drafts);
					if (drafts.length === 0) {
						manageDraftsModal.setAttribute('aria-hidden', 'true');
					} else {
						showManageDrafts(); // Refresh the list
					}
				}
			});
		});
		
		manageDraftsModal.setAttribute('aria-hidden', 'false');
	}

	// Update UI based on draft availability
	function updateDraftUI() {
		const drafts = getAllDrafts();
		const hasDrafts = drafts.length > 0;
		
		if (manageDraftsBtn) {
			manageDraftsBtn.style.display = hasDrafts ? 'inline-block' : 'none';
		}
	}

	// Event listeners
	if (saveDraftBtn) {
		saveDraftBtn.addEventListener('click', saveDraft);
	}

	if (manageDraftsBtn) {
		manageDraftsBtn.addEventListener('click', showManageDrafts);
	}

	if (closeDraftsModal) {
		closeDraftsModal.addEventListener('click', function() {
			manageDraftsModal.setAttribute('aria-hidden', 'true');
		});
	}

	// Close modal when clicking outside
	if (manageDraftsModal) {
		manageDraftsModal.addEventListener('click', function(e) {
			if (e.target === manageDraftsModal) {
				manageDraftsModal.setAttribute('aria-hidden', 'true');
			}
		});
	}

	// Initialize draft UI on page load
	updateDraftUI();

	// Clear the specific loaded draft on successful save (only if a draft was loaded)
	if (shareBtn) {
		shareBtn.addEventListener('click', function() {
			setTimeout(() => {
				// Only prompt to delete if a specific draft was loaded
				if (currentlyLoadedDraftId !== null) {
					const drafts = getAllDrafts();
					const loadedDraft = drafts.find(d => d.id === currentlyLoadedDraftId);
					if (loadedDraft) {
						if (confirm(`Form saved successfully! Would you like to delete the draft "${loadedDraft.name}"?`)) {
							const updatedDrafts = drafts.filter(d => d.id !== currentlyLoadedDraftId);
							localStorage.setItem(DRAFTS_KEY, JSON.stringify(updatedDrafts));
							currentlyLoadedDraftId = null; // Clear the tracking
							updateDraftUI();
						}
					}
				}
				// If no draft was loaded, don't prompt at all
			}, 1000);
		}, true);
	}
});
