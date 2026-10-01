/**
 * NutriPulse - Food & Macronutrient Tracker Script
 * Features: Food Database Search, Auto Goal Calculator (Mifflin-St Jeor),
 * Full 15-Nutrient Breakdown, LocalStorage, Clean Excel-Friendly CSV Export.
 */

// 1. Food Dataset (Nutrients per 100g)
const FOOD_DATABASE = [
  {
    name: "Chicken Breast (Raw)",
    cal: 120, protein: 22.5, fat: 2.6, satFat: 0.7, unsatFat: 1.5, transFat: 0,
    carbs: 0, fiber: 0, sugars: 0, netCarbs: 0, sodium: 74, potassium: 256, cholesterol: 64, calcium: 11, iron: 0.7
  },
  {
    name: "Banana (Fresh)",
    cal: 89, protein: 1.1, fat: 0.3, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 22.8, fiber: 2.6, sugars: 12.2, netCarbs: 20.2, sodium: 1, potassium: 358, cholesterol: 0, calcium: 5, iron: 0.3
  },
  {
    name: "Oats (Rolled)",
    cal: 389, protein: 16.9, fat: 6.9, satFat: 1.2, unsatFat: 5.0, transFat: 0,
    carbs: 66.3, fiber: 10.6, sugars: 0.8, netCarbs: 55.7, sodium: 2, potassium: 429, cholesterol: 0, calcium: 54, iron: 4.7
  },
  {
    name: "Salmon (Fillet)",
    cal: 208, protein: 20.4, fat: 13.4, satFat: 3.1, unsatFat: 8.9, transFat: 0,
    carbs: 0, fiber: 0, sugars: 0, netCarbs: 0, sodium: 59, potassium: 363, cholesterol: 55, calcium: 9, iron: 0.3
  },
  {
    name: "Egg (Whole, Large)",
    cal: 143, protein: 12.6, fat: 9.5, satFat: 3.1, unsatFat: 5.0, transFat: 0,
    carbs: 0.7, fiber: 0, sugars: 0.4, netCarbs: 0.7, sodium: 142, potassium: 138, cholesterol: 372, calcium: 56, iron: 1.8
  },
  {
    name: "White Rice (Cooked)",
    cal: 130, protein: 2.7, fat: 0.3, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 28.2, fiber: 0.4, sugars: 0.1, netCarbs: 27.8, sodium: 1, potassium: 35, cholesterol: 0, calcium: 10, iron: 0.2
  },
  {
    name: "Avocado",
    cal: 160, protein: 2.0, fat: 14.7, satFat: 2.1, unsatFat: 11.5, transFat: 0,
    carbs: 8.5, fiber: 6.7, sugars: 0.7, netCarbs: 1.8, sodium: 7, potassium: 485, cholesterol: 0, calcium: 12, iron: 0.6
  },
  {
    name: "Almonds",
    cal: 579, protein: 21.2, fat: 49.9, satFat: 3.8, unsatFat: 42.1, transFat: 0,
    carbs: 21.6, fiber: 12.5, sugars: 4.4, netCarbs: 9.1, sodium: 1, potassium: 733, cholesterol: 0, calcium: 269, iron: 3.7
  },
  {
    name: "Greek Yogurt (Plain)",
    cal: 59, protein: 10.0, fat: 0.4, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 3.6, fiber: 0, sugars: 3.2, netCarbs: 3.6, sodium: 36, potassium: 141, cholesterol: 5, calcium: 110, iron: 0.1
  },
  {
    name: "Spinach (Raw)",
    cal: 23, protein: 2.9, fat: 0.4, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 3.6, fiber: 2.2, sugars: 0.4, netCarbs: 1.4, sodium: 79, potassium: 558, cholesterol: 0, calcium: 99, iron: 2.7
  }
];

// Default Daily Macro Goals
let dailyTargets = {
  cal: 2000,
  protein: 150,
  carbs: 250,
  fat: 65,
  fiber: 30
};

// Application State Variables
let loggedItems = [];
let currentPreviewItem = null;

// DOM Elements
const calculatorForm = document.getElementById('calculator-form');
const foodInput = document.getElementById('food-input');
const weightInput = document.getElementById('weight-input');
const unitSelect = document.getElementById('unit-select');
const foodForm = document.getElementById('food-form');
const suggestionsBox = document.getElementById('suggestions-box');
const previewSection = document.getElementById('preview-section');
const nutrientGrid = document.getElementById('nutrient-grid');
const previewTitle = document.getElementById('preview-title');
const previewPortion = document.getElementById('preview-portion');
const logPreviewBtn = document.getElementById('log-preview-btn');
const logTableBody = document.getElementById('log-table-body');
const emptyState = document.getElementById('empty-state');
const logCountBadge = document.getElementById('log-count-badge');
const clearAllBtn = document.getElementById('clear-all-btn');
const exportBtn = document.getElementById('export-btn');
const nutrientModal = document.getElementById('nutrient-modal');
const modalTitle = document.getElementById('modal-title');
const modalContent = document.getElementById('modal-content');
const closeModalBtn = document.getElementById('close-modal-btn');

// --- Application Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  loadFromLocalStorage();
  setupEventListeners();
  updateDashboard();
  renderLogTable();
});

// --- Event Listeners Setup ---
function setupEventListeners() {
  calculatorForm.addEventListener('submit', handleTargetCalculation);
  foodInput.addEventListener('input', handleFoodSearch);
  foodForm.addEventListener('submit', handleCalculateSubmit);
  logPreviewBtn.addEventListener('click', () => {
    if (currentPreviewItem) {
      logItem(currentPreviewItem);
      previewSection.classList.add('hidden');
    }
  });
  clearAllBtn.addEventListener('click', clearAllLog);
  exportBtn.addEventListener('click', exportCleanCSV);
  closeModalBtn.addEventListener('click', closeModal);

  document.addEventListener('click', (e) => {
    if (!foodInput.contains(e.target) && !suggestionsBox.contains(e.target)) {
      suggestionsBox.classList.add('hidden');
    }
  });
}

// --- NEW FEATURE 1: Goal & Macro Calculator (Mifflin-St Jeor Formula) ---
function handleTargetCalculation(e) {
  e.preventDefault();
  
  const age = parseInt(document.getElementById('calc-age').value);
  const gender = document.getElementById('calc-gender').value;
  const weight = parseFloat(document.getElementById('calc-weight').value);
  const height = parseFloat(document.getElementById('calc-height').value);
  const activity = parseFloat(document.getElementById('calc-activity').value);
  const goal = document.getElementById('calc-goal').value;

  if (isNaN(age) || isNaN(weight) || isNaN(height)) return;

  // 1. Calculate Basal Metabolic Rate (BMR)
  let bmr = (10 * weight) + (6.25 * height) - (5 * age);
  bmr += (gender === 'male') ? 5 : -161;

  // 2. Total Daily Energy Expenditure (TDEE)
  let tdee = bmr * activity;

  // 3. Adjust for Goal
  if (goal === 'lose') tdee *= 0.85; // 15% deficit
  if (goal === 'gain') tdee *= 1.15; // 15% surplus

  const targetCal = Math.round(tdee);
  
  // 4. Macro Splits: Protein (2g per kg), Fat (25% total cal), Carbs (Remainder)
  const targetProtein = Math.round(weight * 2.0);
  const targetFat = Math.round((targetCal * 0.25) / 9);
  const remainingCalForCarbs = targetCal - (targetProtein * 4) - (targetFat * 9);
  const targetCarbs = Math.max(50, Math.round(remainingCalForCarbs / 4));
  const targetFiber = Math.round((targetCal / 1000) * 14); // 14g fiber per 1000 kcal

  dailyTargets = {
    cal: targetCal,
    protein: targetProtein,
    carbs: targetCarbs,
    fat: targetFat,
    fiber: targetFiber
  };

  // Update UI & LocalStorage
  document.getElementById('target-mode-badge').textContent = `Personalized Goal (${goal.toUpperCase()})`;
  saveToLocalStorage();
  updateDashboard();
  
  alert(`Goals Calculated & Updated successfully!\nDaily Calories: ${targetCal} kcal\nProtein: ${targetProtein}g | Carbs: ${targetCarbs}g | Fat: ${targetFat}g`);
}

// --- Food Auto-suggest Search ---
function handleFoodSearch() {
  const query = foodInput.value.toLowerCase().trim();
  if (!query) {
    suggestionsBox.classList.add('hidden');
    return;
  }

  const matches = FOOD_DATABASE.filter(item => item.name.toLowerCase().includes(query));

  if (matches.length > 0) {
    suggestionsBox.innerHTML = matches.map(item => `
      <div class="suggestion-item" onclick="selectSuggestion('${item.name}')">${item.name}</div>
    `).join('');
    suggestionsBox.classList.remove('hidden');
  } else {
    suggestionsBox.classList.add('hidden');
  }
}

function selectSuggestion(name) {
  foodInput.value = name;
  suggestionsBox.classList.add('hidden');
}

// --- Core Nutrient Scaling Logic ---
function calculateNutrients(baseItem, weightVal, unitVal) {
  const weightInGrams = unitVal === 'oz' ? weightVal * 28.3495 : weightVal;
  const factor = weightInGrams / 100;

  return {
    id: Date.now(),
    name: baseItem.name,
    weightEntered: weightVal,
    unit: unitVal,
    weightInGrams: Math.round(weightInGrams),
    cal: Math.round(baseItem.cal * factor),
    protein: +(baseItem.protein * factor).toFixed(1),
    fat: +(baseItem.fat * factor).toFixed(1),
    satFat: +(baseItem.satFat * factor).toFixed(1),
    unsatFat: +(baseItem.unsatFat * factor).toFixed(1),
    transFat: +(baseItem.transFat * factor).toFixed(1),
    carbs: +(baseItem.carbs * factor).toFixed(1),
    fiber: +(baseItem.fiber * factor).toFixed(1),
    sugars: +(baseItem.sugars * factor).toFixed(1),
    netCarbs: +(baseItem.netCarbs * factor).toFixed(1),
    sodium: Math.round(baseItem.sodium * factor),
    potassium: Math.round(baseItem.potassium * factor),
    cholesterol: Math.round(baseItem.cholesterol * factor),
    calcium: Math.round(baseItem.calcium * factor),
    iron: +(baseItem.iron * factor).toFixed(1)
  };
}

function handleCalculateSubmit(e) {
  e.preventDefault();
  const searchName = foodInput.value.trim();
  const weight = parseFloat(weightInput.value);
  const unit = unitSelect.value;

  if (!searchName || isNaN(weight) || weight <= 0) return;

  let matchedItem = FOOD_DATABASE.find(item => item.name.toLowerCase() === searchName.toLowerCase());

  if (!matchedItem) {
    matchedItem = {
      name: searchName,
      cal: 150, protein: 5.0, fat: 3.0, satFat: 1.0, unsatFat: 1.5, transFat: 0,
      carbs: 20.0, fiber: 2.0, sugars: 5.0, netCarbs: 18.0, sodium: 100, potassium: 200, cholesterol: 10, calcium: 20, iron: 1.0
    };
  }

  currentPreviewItem = calculateNutrients(matchedItem, weight, unit);
  displayPreviewCard(currentPreviewItem);
}

// --- Preview Breakdown Display ---
function displayPreviewCard(item) {
  previewTitle.textContent = item.name;
  previewPortion.textContent = `${item.weightEntered} ${item.unit} (${item.weightInGrams}g)`;

  const metrics = [
    { label: "Total Calories", val: `${item.cal} kcal` },
    { label: "Protein", val: `${item.protein} g` },
    { label: "Total Fat", val: `${item.fat} g` },
    { label: "Saturated Fat", val: `${item.satFat} g` },
    { label: "Unsaturated Fat", val: `${item.unsatFat} g` },
    { label: "Trans Fat", val: `${item.transFat} g` },
    { label: "Total Carbs", val: `${item.carbs} g` },
    { label: "Dietary Fiber", val: `${item.fiber} g` },
    { label: "Total Sugars", val: `${item.sugars} g` },
    { label: "Net Carbs", val: `${item.netCarbs} g` },
    { label: "Sodium", val: `${item.sodium} mg` },
    { label: "Potassium", val: `${item.potassium} mg` },
    { label: "Cholesterol", val: `${item.cholesterol} mg` },
    { label: "Calcium", val: `${item.calcium} mg` },
    { label: "Iron", val: `${item.iron} mg` }
  ];

  nutrientGrid.innerHTML = metrics.map(m => `
    <div class="nutrient-chip">
      <span class="name">${m.label}</span>
      <span class="value">${m.val}</span>
    </div>
  `).join('');

  previewSection.classList.remove('hidden');
  previewSection.scrollIntoView({ behavior: 'smooth' });
}

// --- Log & State Management ---
function logItem(item) {
  loggedItems.push(item);
  saveToLocalStorage();
  updateDashboard();
  renderLogTable();
}

function deleteItem(id) {
  loggedItems = loggedItems.filter(item => item.id !== id);
  saveToLocalStorage();
  updateDashboard();
  renderLogTable();
}

function clearAllLog() {
  if (loggedItems.length === 0) return;
  if (confirm("Are you sure you want to clear today's entire food log?")) {
    loggedItems = [];
    saveToLocalStorage();
    updateDashboard();
    renderLogTable();
  }
}

// --- Dashboard Summary Calculations ---
function updateDashboard() {
  const totals = loggedItems.reduce((acc, curr) => ({
    cal: acc.cal + curr.cal,
    protein: acc.protein + curr.protein,
    carbs: acc.carbs + curr.carbs,
    fat: acc.fat + curr.fat,
    fiber: acc.fiber + curr.fiber
  }), { cal: 0, protein: 0, carbs: 0, fat: 0, fiber: 0 });

  // Render Daily Totals
  document.getElementById('summary-cal').innerHTML = `${totals.cal.toLocaleString()} <small>kcal</small>`;
  document.getElementById('summary-protein').innerHTML = `${totals.protein.toFixed(1)} <small>g</small>`;
  document.getElementById('summary-carbs').innerHTML = `${totals.carbs.toFixed(1)} <small>g</small>`;
  document.getElementById('summary-fat').innerHTML = `${totals.fat.toFixed(1)} <small>g</small>`;
  document.getElementById('summary-fiber').innerHTML = `${totals.fiber.toFixed(1)} <small>g</small>`;

  // Update Dynamic Target Labels
  document.getElementById('target-text-cal').textContent = `Target: ${dailyTargets.cal.toLocaleString()} kcal`;
  document.getElementById('target-text-protein').textContent = `Target: ${dailyTargets.protein}g`;
  document.getElementById('target-text-carbs').textContent = `Target: ${dailyTargets.carbs}g`;
  document.getElementById('target-text-fat').textContent = `Target: ${dailyTargets.fat}g`;
  document.getElementById('target-text-fiber').textContent = `Target: ${dailyTargets.fiber}g`;

  // Update Progress Bars
  document.getElementById('cal-progress').style.width = `${Math.min((totals.cal / dailyTargets.cal) * 100, 100)}%`;
  document.getElementById('protein-progress').style.width = `${Math.min((totals.protein / dailyTargets.protein) * 100, 100)}%`;
  document.getElementById('carbs-progress').style.width = `${Math.min((totals.carbs / dailyTargets.carbs) * 100, 100)}%`;
  document.getElementById('fat-progress').style.width = `${Math.min((totals.fat / dailyTargets.fat) * 100, 100)}%`;
  document.getElementById('fiber-progress').style.width = `${Math.min((totals.fiber / dailyTargets.fiber) * 100, 100)}%`;
}

// --- Render Table List ---
function renderLogTable() {
  logCountBadge.textContent = `${loggedItems.length} Item${loggedItems.length !== 1 ? 's' : ''}`;

  if (loggedItems.length === 0) {
    emptyState.classList.remove('hidden');
    logTableBody.innerHTML = '';
    return;
  }

  emptyState.classList.add('hidden');
  logTableBody.innerHTML = loggedItems.map(item => `
    <tr onclick="openNutrientModal(${item.id})">
      <td><strong>${item.name}</strong></td>
      <td>${item.weightEntered} ${item.unit}</td>
      <td>${item.cal} kcal</td>
      <td>${item.protein} g</td>
      <td>${item.fat} g</td>
      <td>${item.carbs} g</td>
      <td>${item.fiber} g</td>
      <td class="text-right" onclick="event.stopPropagation()">
        <button onclick="deleteItem(${item.id})" class="btn-icon" title="Delete entry">
          <i class="fa-solid fa-trash"></i>
        </button>
      </td>
    </tr>
  `).join('');
}

// --- Modal Popup View ---
function openNutrientModal(id) {
  const item = loggedItems.find(i => i.id === id);
  if (!item) return;

  modalTitle.textContent = `${item.name} (${item.weightEntered} ${item.unit})`;
  
  const metrics = [
    { label: "Calories", val: `${item.cal} kcal` },
    { label: "Protein", val: `${item.protein} g` },
    { label: "Total Fat", val: `${item.fat} g` },
    { label: "Saturated Fat", val: `${item.satFat} g` },
    { label: "Unsaturated Fat", val: `${item.unsatFat} g` },
    { label: "Trans Fat", val: `${item.transFat} g` },
    { label: "Total Carbohydrates", val: `${item.carbs} g` },
    { label: "Dietary Fiber", val: `${item.fiber} g` },
    { label: "Total Sugars", val: `${item.sugars} g` },
    { label: "Net Carbs", val: `${item.netCarbs} g` },
    { label: "Sodium", val: `${item.sodium} mg` },
    { label: "Potassium", val: `${item.potassium} mg` },
    { label: "Cholesterol", val: `${item.cholesterol} mg` },
    { label: "Calcium", val: `${item.calcium} mg` },
    { label: "Iron", val: `${item.iron} mg` }
  ];

  modalContent.innerHTML = `
    <div class="nutrient-grid">
      ${metrics.map(m => `
        <div class="nutrient-chip">
          <span class="name">${m.label}</span>
          <span class="value">${m.val}</span>
        </div>
      `).join('')}
    </div>
  `;

  nutrientModal.classList.remove('hidden');
}

function closeModal() {
  nutrientModal.classList.add('hidden');
}

// --- NEW FEATURE 2: Clean, Properly Standardized CSV Export ---
function exportCleanCSV() {
  if (loggedItems.length === 0) {
    alert("No logged food items to export!");
    return;
  }

  const currentDate = new Date().toLocaleDateString();
  
  // 1. Calculate Daily Totals
  const totals = loggedItems.reduce((acc, curr) => ({
    cal: acc.cal + curr.cal,
    protein: acc.protein + curr.protein,
    fat: acc.fat + curr.fat,
    satFat: acc.satFat + curr.satFat,
    carbs: acc.carbs + curr.carbs,
    fiber: acc.fiber + curr.fiber,
    sugars: acc.sugars + curr.sugars,
    sodium: acc.sodium + curr.sodium,
    potassium: acc.potassium + curr.potassium,
    cholesterol: acc.cholesterol + curr.cholesterol,
  }), { cal: 0, protein: 0, fat: 0, satFat: 0, carbs: 0, fiber: 0, sugars: 0, sodium: 0, potassium: 0, cholesterol: 0 });

  // 2. Format Professional Header Summary & Column Titles
  let csvRows = [
    `"NUTRIPULSE DAILY NUTRITION LOG"`,
    `"Date Exported:","${currentDate}"`,
    `"Daily Calorie Target:","${dailyTargets.cal} kcal"`,
    `"Daily Protein Target:","${dailyTargets.protein} g"`,
    `""`, // Blank Separator Line
    [
      "Food Name",
      "Quantity Entered",
      "Weight (g)",
      "Calories (kcal)",
      "Protein (g)",
      "Total Fat (g)",
      "Saturated Fat (g)",
      "Carbohydrates (g)",
      "Dietary Fiber (g)",
      "Sugars (g)",
      "Sodium (mg)",
      "Potassium (mg)",
      "Cholesterol (mg)"
    ].map(header => `"${header}"`).join(",")
  ];

  // 3. Add Individual Logged Rows
  loggedItems.forEach(item => {
    const row = [
      `"${item.name.replace(/"/g, '""')}"`,
      `"${item.weightEntered} ${item.unit}"`,
      item.weightInGrams,
      item.cal,
      item.protein,
      item.fat,
      item.satFat,
      item.carbs,
      item.fiber,
      item.sugars,
      item.sodium,
      item.potassium,
      item.cholesterol
    ];
    csvRows.push(row.join(","));
  });

  // 4. Add Clear Summary Footer Row
  csvRows.push(`""`); // Blank line
  const summaryRow = [
    `"TOTAL DAILY CONSUMPTION"`,
    `"-"`,
    `"-"`,
    totals.cal,
    totals.protein.toFixed(1),
    totals.fat.toFixed(1),
    totals.satFat.toFixed(1),
    totals.carbs.toFixed(1),
    totals.fiber.toFixed(1),
    totals.sugars.toFixed(1),
    totals.sodium,
    totals.potassium,
    totals.cholesterol
  ];
  csvRows.push(summaryRow.join(","));

  // 5. Trigger File Download with Universal BOM for MS Excel Compatibility
  const csvString = "\uFEFF" + csvRows.join("\n");
  const blob = new Blob([csvString], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const downloadLink = document.createElement("a");
  downloadLink.href = url;
  downloadLink.setAttribute("download", `NutriPulse_Log_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(downloadLink);
  downloadLink.click();
  document.body.removeChild(downloadLink);
}

// --- LocalStorage Logic ---
function saveToLocalStorage() {
  localStorage.setItem('nutripulse_food_log', JSON.stringify(loggedItems));
  localStorage.setItem('nutripulse_daily_targets', JSON.stringify(dailyTargets));
}

function loadFromLocalStorage() {
  const savedLog = localStorage.getItem('nutripulse_food_log');
  if (savedLog) {
    try { loggedItems = JSON.parse(savedLog); } catch (e) { loggedItems = []; }
  }

  const savedTargets = localStorage.getItem('nutripulse_daily_targets');
  if (savedTargets) {
    try { 
      dailyTargets = JSON.parse(savedTargets); 
      document.getElementById('target-mode-badge').textContent = 'Custom Goals Loaded';
    } catch (e) {}
  }
}