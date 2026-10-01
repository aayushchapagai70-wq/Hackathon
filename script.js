/**
 * NutriPulse - Food & Macronutrient Tracker Script
 * Features: 100+ Nepalese Food Database, Auto Goal Calculator (Mifflin-St Jeor),
 * Full 15-Nutrient Breakdown, LocalStorage, Clean Excel-Friendly CSV Export.
 */

// 1. Comprehensive Nepalese Food Dataset (Nutrients per 100g)
const FOOD_DATABASE = [
  // --- STAPLES & GRAINS (Bhat, Roti, Dhindo, Sel Roti) ---
  {
    name: "Steamed White Rice (Bhat)",
    cal: 130, protein: 2.7, fat: 0.3, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 28.2, fiber: 0.4, sugars: 0.1, netCarbs: 27.8, sodium: 1, potassium: 35, cholesterol: 0, calcium: 10, iron: 0.2
  },
  {
    name: "Brown Rice (Cooked)",
    cal: 111, protein: 2.6, fat: 0.9, satFat: 0.2, unsatFat: 0.6, transFat: 0,
    carbs: 23.0, fiber: 1.8, sugars: 0.4, netCarbs: 21.2, sodium: 5, potassium: 86, cholesterol: 0, calcium: 10, iron: 0.5
  },
  {
    name: "Wheat Roti / Chapati (Plain)",
    cal: 297, protein: 9.2, fat: 3.5, satFat: 0.6, unsatFat: 2.5, transFat: 0,
    carbs: 56.0, fiber: 9.0, sugars: 1.0, netCarbs: 47.0, sodium: 190, potassium: 250, cholesterol: 0, calcium: 40, iron: 3.0
  },
  {
    name: "Millet Dhindo (Kodo ko Dhindo)",
    cal: 115, protein: 2.8, fat: 0.7, satFat: 0.2, unsatFat: 0.4, transFat: 0,
    carbs: 24.0, fiber: 3.2, sugars: 0.2, netCarbs: 20.8, sodium: 5, potassium: 120, cholesterol: 0, calcium: 35, iron: 1.8
  },
  {
    name: "Buckwheat Dhindo (Fapar ko Dhindo)",
    cal: 120, protein: 3.5, fat: 0.9, satFat: 0.2, unsatFat: 0.6, transFat: 0,
    carbs: 25.0, fiber: 3.8, sugars: 0.5, netCarbs: 21.2, sodium: 4, potassium: 150, cholesterol: 0, calcium: 18, iron: 1.5
  },
  {
    name: "Corn Dhindo (Makai ko Dhindo)",
    cal: 108, protein: 2.5, fat: 0.8, satFat: 0.1, unsatFat: 0.6, transFat: 0,
    carbs: 23.5, fiber: 2.8, sugars: 0.6, netCarbs: 20.7, sodium: 3, potassium: 110, cholesterol: 0, calcium: 12, iron: 1.1
  },
  {
    name: "Sel Roti",
    cal: 380, protein: 5.2, fat: 14.5, satFat: 3.2, unsatFat: 10.5, transFat: 0,
    carbs: 58.0, fiber: 1.5, sugars: 16.0, netCarbs: 56.5, sodium: 45, potassium: 95, cholesterol: 0, calcium: 22, iron: 1.2
  },
  {
    name: "Puri (Fried Wheat Bread)",
    cal: 330, protein: 6.8, fat: 17.5, satFat: 4.2, unsatFat: 12.5, transFat: 0,
    carbs: 37.0, fiber: 2.5, sugars: 0.8, netCarbs: 34.5, sodium: 220, potassium: 140, cholesterol: 0, calcium: 25, iron: 2.1
  },
  {
    name: "Beaten Rice (Bhoja / Chiura - Dry)",
    cal: 346, protein: 6.6, fat: 1.2, satFat: 0.3, unsatFat: 0.8, transFat: 0,
    carbs: 77.3, fiber: 2.1, sugars: 0.5, netCarbs: 75.2, sodium: 12, potassium: 160, cholesterol: 0, calcium: 20, iron: 5.5
  },
  {
    name: "Roasted Corn (Bhutta)",
    cal: 130, protein: 3.4, fat: 1.5, satFat: 0.2, unsatFat: 1.1, transFat: 0,
    carbs: 27.0, fiber: 3.0, sugars: 3.2, netCarbs: 24.0, sodium: 15, potassium: 270, cholesterol: 0, calcium: 7, iron: 0.9
  },
  {
    name: "Popcorn (Makai ko Khaja - Plain)",
    cal: 387, protein: 12.9, fat: 4.5, satFat: 0.6, unsatFat: 3.5, transFat: 0,
    carbs: 77.9, fiber: 14.5, sugars: 0.9, netCarbs: 63.4, sodium: 8, potassium: 329, cholesterol: 0, calcium: 7, iron: 3.2
  },
  {
    name: "Naan (Plain)",
    cal: 290, protein: 8.5, fat: 5.0, satFat: 1.2, unsatFat: 3.5, transFat: 0,
    carbs: 52.0, fiber: 2.2, sugars: 3.0, netCarbs: 49.8, sodium: 420, potassium: 120, cholesterol: 5, calcium: 45, iron: 2.8
  },
  {
    name: "Bhatura",
    cal: 310, protein: 7.2, fat: 12.0, satFat: 2.8, unsatFat: 8.5, transFat: 0,
    carbs: 43.0, fiber: 1.8, sugars: 2.0, netCarbs: 41.2, sodium: 380, potassium: 110, cholesterol: 0, calcium: 30, iron: 2.0
  },
  {
    name: "Oats (Rolled - Cooked in Water)",
    cal: 71, protein: 2.5, fat: 1.4, satFat: 0.2, unsatFat: 1.0, transFat: 0,
    carbs: 12.0, fiber: 1.7, sugars: 0.3, netCarbs: 10.3, sodium: 2, potassium: 61, cholesterol: 0, calcium: 9, iron: 0.9
  },

  // --- LENTILS & BEANS (Dal & Kwati) ---
  {
    name: "Yellow Lentils Cooked (Rahu / Toor Dal)",
    cal: 116, protein: 6.8, fat: 1.5, satFat: 0.3, unsatFat: 1.0, transFat: 0,
    carbs: 20.0, fiber: 3.8, sugars: 1.2, netCarbs: 16.2, sodium: 180, potassium: 240, cholesterol: 0, calcium: 22, iron: 1.6
  },
  {
    name: "Black Lentils Cooked (Kalo Dal / Kalo Maas)",
    cal: 110, protein: 7.5, fat: 0.8, satFat: 0.2, unsatFat: 0.5, transFat: 0,
    carbs: 19.2, fiber: 4.5, sugars: 0.8, netCarbs: 14.7, sodium: 160, potassium: 280, cholesterol: 0, calcium: 45, iron: 2.3
  },
  {
    name: "Red Lentils Cooked (Musuro ko Dal)",
    cal: 116, protein: 9.0, fat: 0.4, satFat: 0.1, unsatFat: 0.2, transFat: 0,
    carbs: 20.0, fiber: 7.9, sugars: 1.8, netCarbs: 12.1, sodium: 170, potassium: 369, cholesterol: 0, calcium: 19, iron: 3.3
  },
  {
    name: "Mung Dal Cooked (Moong)",
    cal: 105, protein: 7.0, fat: 0.6, satFat: 0.1, unsatFat: 0.4, transFat: 0,
    carbs: 19.0, fiber: 5.2, sugars: 1.5, netCarbs: 13.8, sodium: 150, potassium: 290, cholesterol: 0, calcium: 27, iron: 1.4
  },
  {
    name: "Kwati (Sprouted Mixed Beans Soup)",
    cal: 135, protein: 8.5, fat: 2.1, satFat: 0.4, unsatFat: 1.5, transFat: 0,
    carbs: 21.0, fiber: 6.2, sugars: 2.0, netCarbs: 14.8, sodium: 210, potassium: 380, cholesterol: 0, calcium: 52, iron: 2.9
  },
  {
    name: "Boiled Chickpeas (Chana)",
    cal: 164, protein: 8.9, fat: 2.6, satFat: 0.3, unsatFat: 2.1, transFat: 0,
    carbs: 27.4, fiber: 7.6, sugars: 4.8, netCarbs: 19.8, sodium: 24, potassium: 291, cholesterol: 0, calcium: 49, iron: 2.9
  },
  {
    name: "Rajma Curry (Kidney Beans)",
    cal: 140, protein: 7.2, fat: 4.1, satFat: 0.6, unsatFat: 3.2, transFat: 0,
    carbs: 19.5, fiber: 5.5, sugars: 2.1, netCarbs: 14.0, sodium: 290, potassium: 340, cholesterol: 0, calcium: 38, iron: 2.2
  },
  {
    name: "Gundruk Curry (Fermented Leafy Soup)",
    cal: 45, protein: 3.8, fat: 1.2, satFat: 0.2, unsatFat: 0.8, transFat: 0,
    carbs: 5.5, fiber: 3.1, sugars: 0.5, netCarbs: 2.4, sodium: 310, potassium: 210, cholesterol: 0, calcium: 110, iron: 4.2
  },

  // --- POPULAR STREET FOOD & SNACKS (Momo, Chowmein, Laphing) ---
  {
    name: "Steam Chicken Momo",
    cal: 185, protein: 11.2, fat: 6.5, satFat: 1.8, unsatFat: 4.2, transFat: 0,
    carbs: 20.5, fiber: 1.2, sugars: 1.0, netCarbs: 19.3, sodium: 340, potassium: 180, cholesterol: 32, calcium: 18, iron: 1.4
  },
  {
    name: "Steam Buff Momo",
    cal: 195, protein: 12.5, fat: 7.2, satFat: 2.5, unsatFat: 4.2, transFat: 0,
    carbs: 20.0, fiber: 1.2, sugars: 0.9, netCarbs: 18.8, sodium: 360, potassium: 195, cholesterol: 38, calcium: 20, iron: 2.1
  },
  {
    name: "Steam Veg Momo",
    cal: 150, protein: 4.5, fat: 4.8, satFat: 0.8, unsatFat: 3.8, transFat: 0,
    carbs: 22.0, fiber: 2.1, sugars: 1.5, netCarbs: 19.9, sodium: 290, potassium: 140, cholesterol: 0, calcium: 25, iron: 1.1
  },
  {
    name: "Fried Chicken Momo",
    cal: 260, protein: 10.8, fat: 14.2, satFat: 3.5, unsatFat: 10.1, transFat: 0,
    carbs: 22.0, fiber: 1.2, sugars: 1.0, netCarbs: 20.8, sodium: 390, potassium: 170, cholesterol: 35, calcium: 19, iron: 1.3
  },
  {
    name: "Chicken C-Momo (Spicy)",
    cal: 215, protein: 10.5, fat: 9.8, satFat: 2.4, unsatFat: 7.0, transFat: 0,
    carbs: 21.0, fiber: 1.5, sugars: 3.2, netCarbs: 19.5, sodium: 520, potassium: 210, cholesterol: 30, calcium: 22, iron: 1.6
  },
  {
    name: "Chicken Chowmein",
    cal: 210, protein: 9.5, fat: 8.2, satFat: 1.6, unsatFat: 6.2, transFat: 0,
    carbs: 24.5, fiber: 1.8, sugars: 2.0, netCarbs: 22.7, sodium: 480, potassium: 190, cholesterol: 25, calcium: 22, iron: 1.5
  },
  {
    name: "Veg Chowmein",
    cal: 175, protein: 4.2, fat: 6.8, satFat: 1.1, unsatFat: 5.4, transFat: 0,
    carbs: 24.0, fiber: 2.3, sugars: 2.2, netCarbs: 21.7, sodium: 410, potassium: 150, cholesterol: 0, calcium: 28, iron: 1.2
  },
  {
    name: "Buff Chowmein",
    cal: 225, protein: 10.8, fat: 9.1, satFat: 2.3, unsatFat: 6.3, transFat: 0,
    carbs: 24.0, fiber: 1.8, sugars: 1.9, netCarbs: 22.2, sodium: 510, potassium: 205, cholesterol: 32, calcium: 24, iron: 2.2
  },
  {
    name: "Laphing (Cold Noodle - Dry/Soup)",
    cal: 140, protein: 3.8, fat: 4.2, satFat: 0.6, unsatFat: 3.4, transFat: 0,
    carbs: 22.0, fiber: 1.1, sugars: 1.0, netCarbs: 20.9, sodium: 620, potassium: 90, cholesterol: 0, calcium: 15, iron: 0.8
  },
  {
    name: "Chatpate",
    cal: 190, protein: 5.2, fat: 7.5, satFat: 1.2, unsatFat: 5.8, transFat: 0,
    carbs: 26.0, fiber: 3.5, sugars: 2.8, netCarbs: 22.5, sodium: 580, potassium: 230, cholesterol: 0, calcium: 32, iron: 2.8
  },
  {
    name: "Pani Puri (6 pcs approx. filled)",
    cal: 160, protein: 3.1, fat: 6.2, satFat: 1.1, unsatFat: 4.8, transFat: 0,
    carbs: 23.0, fiber: 2.2, sugars: 1.5, netCarbs: 20.8, sodium: 450, potassium: 140, cholesterol: 0, calcium: 18, iron: 1.2
  },
  {
    name: "Samosa (1 pc medium ~100g)",
    cal: 262, protein: 4.5, fat: 12.8, satFat: 3.2, unsatFat: 9.1, transFat: 0,
    carbs: 32.2, fiber: 2.8, sugars: 1.8, netCarbs: 29.4, sodium: 422, potassium: 215, cholesterol: 0, calcium: 24, iron: 1.6
  },
  {
    name: "Samosa Chaat",
    cal: 195, protein: 5.1, fat: 8.5, satFat: 2.1, unsatFat: 6.0, transFat: 0,
    carbs: 25.0, fiber: 3.2, sugars: 4.5, netCarbs: 21.8, sodium: 490, potassium: 260, cholesterol: 5, calcium: 45, iron: 1.8
  },
  {
    name: "Aloo Chop",
    cal: 210, protein: 3.8, fat: 10.5, satFat: 2.2, unsatFat: 7.8, transFat: 0,
    carbs: 25.0, fiber: 2.5, sugars: 1.2, netCarbs: 22.5, sodium: 380, potassium: 310, cholesterol: 0, calcium: 18, iron: 1.4
  },
  {
    name: "Pakoda (Vegetable Onion)",
    cal: 270, protein: 6.2, fat: 16.0, satFat: 2.8, unsatFat: 12.5, transFat: 0,
    carbs: 25.0, fiber: 4.1, sugars: 2.5, netCarbs: 20.9, sodium: 340, potassium: 280, cholesterol: 0, calcium: 35, iron: 2.1
  },
  {
    name: "Thukpa (Chicken Noodle Soup)",
    cal: 125, protein: 7.8, fat: 3.5, satFat: 0.9, unsatFat: 2.4, transFat: 0,
    carbs: 15.5, fiber: 1.2, sugars: 1.1, netCarbs: 14.3, sodium: 460, potassium: 160, cholesterol: 22, calcium: 18, iron: 1.1
  },
  {
    name: "Bara / Wo (Lentil Patty - Plain)",
    cal: 180, protein: 9.8, fat: 6.2, satFat: 0.9, unsatFat: 5.0, transFat: 0,
    carbs: 21.5, fiber: 4.8, sugars: 1.0, netCarbs: 16.7, sodium: 280, potassium: 320, cholesterol: 0, calcium: 42, iron: 2.6
  },
  {
    name: "Egg Bara / Wo",
    cal: 210, protein: 12.2, fat: 9.8, satFat: 2.2, unsatFat: 7.0, transFat: 0,
    carbs: 18.0, fiber: 3.8, sugars: 0.9, netCarbs: 14.2, sodium: 330, potassium: 310, cholesterol: 160, calcium: 55, iron: 3.0
  },
  {
    name: "Yomari (Chaku Filled)",
    cal: 295, protein: 3.8, fat: 4.2, satFat: 1.5, unsatFat: 2.4, transFat: 0,
    carbs: 61.0, fiber: 2.1, sugars: 28.0, netCarbs: 58.9, sodium: 30, potassium: 180, cholesterol: 0, calcium: 48, iron: 2.1
  },
  {
    name: "Yomari (Khuwa Filled)",
    cal: 315, protein: 6.2, fat: 8.5, satFat: 4.8, unsatFat: 3.2, transFat: 0,
    carbs: 54.0, fiber: 1.2, sugars: 24.0, netCarbs: 52.8, sodium: 65, potassium: 210, cholesterol: 22, calcium: 130, iron: 1.2
  },

  // --- MEAT, POULTRY, FISH & SUKUTI ---
  {
    name: "Chicken Curry (Nepalese Style)",
    cal: 165, protein: 16.5, fat: 9.5, satFat: 2.1, unsatFat: 6.8, transFat: 0,
    carbs: 3.2, fiber: 0.8, sugars: 0.9, netCarbs: 2.4, sodium: 380, potassium: 270, cholesterol: 58, calcium: 22, iron: 1.2
  },
  {
    name: "Chicken Choila",
    cal: 195, protein: 18.2, fat: 11.5, satFat: 2.2, unsatFat: 8.8, transFat: 0,
    carbs: 4.5, fiber: 1.2, sugars: 1.0, netCarbs: 3.3, sodium: 480, potassium: 290, cholesterol: 62, calcium: 25, iron: 1.6
  },
  {
    name: "Chicken Roast / Fried",
    cal: 240, protein: 21.0, fat: 15.5, satFat: 3.8, unsatFat: 11.0, transFat: 0,
    carbs: 3.0, fiber: 0.3, sugars: 0.2, netCarbs: 2.7, sodium: 410, potassium: 260, cholesterol: 75, calcium: 18, iron: 1.3
  },
  {
    name: "Mutton / Goat Curry (Khasi ko Masu)",
    cal: 210, protein: 18.0, fat: 14.0, satFat: 5.2, unsatFat: 8.0, transFat: 0,
    carbs: 3.0, fiber: 0.7, sugars: 0.8, netCarbs: 2.3, sodium: 390, potassium: 310, cholesterol: 68, calcium: 19, iron: 2.8
  },
  {
    name: "Goat Bhutan (Khasi ko Bhutan)",
    cal: 235, protein: 16.8, fat: 17.2, satFat: 6.1, unsatFat: 10.2, transFat: 0,
    carbs: 2.5, fiber: 0.4, sugars: 0.5, netCarbs: 2.1, sodium: 440, potassium: 240, cholesterol: 180, calcium: 22, iron: 4.5
  },
  {
    name: "Buff Curry (Ranga ko Masu)",
    cal: 175, protein: 20.5, fat: 9.0, satFat: 3.1, unsatFat: 5.4, transFat: 0,
    carbs: 2.8, fiber: 0.6, sugars: 0.7, netCarbs: 2.2, sodium: 410, potassium: 320, cholesterol: 55, calcium: 18, iron: 3.2
  },

  {
    name: "Buff Choila",
    cal: 205, protein: 21.5, fat: 11.8, satFat: 3.8, unsatFat: 7.4, transFat: 0,
    carbs: 3.8, fiber: 1.0, sugars: 0.8, netCarbs: 2.8, sodium: 510, potassium: 340, cholesterol: 60, calcium: 22, iron: 3.8
  },
  {
    name: "Buff Sukuti (Dried Meat)",
    cal: 290, protein: 42.0, fat: 12.5, satFat: 4.5, unsatFat: 7.2, transFat: 0,
    carbs: 2.0, fiber: 0.3, sugars: 0.3, netCarbs: 1.7, sodium: 780, potassium: 490, cholesterol: 95, calcium: 28, iron: 6.5
  },
  {
    name: "Pork Curry (Sungur ko Masu)",
    cal: 255, protein: 17.5, fat: 19.5, satFat: 7.1, unsatFat: 11.8, transFat: 0,
    carbs: 2.5, fiber: 0.5, sugars: 0.6, netCarbs: 2.0, sodium: 370, potassium: 280, cholesterol: 65, calcium: 16, iron: 1.4
  },
  {
    name: "Pork Sekuwa",
    cal: 275, protein: 19.2, fat: 21.0, satFat: 7.8, unsatFat: 12.2, transFat: 0,
    carbs: 2.2, fiber: 0.4, sugars: 0.5, netCarbs: 1.8, sodium: 460, potassium: 290, cholesterol: 70, calcium: 18, iron: 1.5
  },
  {
    name: "Local Fish Curry (Machha ko Jhol)",
    cal: 145, protein: 16.8, fat: 7.5, satFat: 1.4, unsatFat: 5.8, transFat: 0,
    carbs: 2.8, fiber: 0.6, sugars: 0.7, netCarbs: 2.2, sodium: 360, potassium: 310, cholesterol: 52, calcium: 35, iron: 1.1
  },
  {
    name: "Fried Fish (Tara ko Machha)",
    cal: 215, protein: 18.5, fat: 14.5, satFat: 2.5, unsatFat: 11.5, transFat: 0,
    carbs: 2.5, fiber: 0.2, sugars: 0.1, netCarbs: 2.3, sodium: 390, potassium: 290, cholesterol: 58, calcium: 40, iron: 1.0
  },
  {
    name: "Egg Curry (Aanda Curry)",
    cal: 155, protein: 10.5, fat: 11.2, satFat: 3.2, unsatFat: 7.2, transFat: 0,
    carbs: 3.2, fiber: 0.6, sugars: 1.1, netCarbs: 2.6, sodium: 380, potassium: 170, cholesterol: 210, calcium: 52, iron: 1.8
  },
  {
    name: "Boiled Egg (1 Large ~50g / values for 100g)",
    cal: 155, protein: 12.6, fat: 10.6, satFat: 3.3, unsatFat: 6.3, transFat: 0,
    carbs: 1.1, fiber: 0, sugars: 1.1, netCarbs: 1.1, sodium: 124, potassium: 126, cholesterol: 373, calcium: 50, iron: 1.2
  },
  {
    name: "Omelette (Nepalese Style with Onion/Chilly)",
    cal: 185, protein: 11.8, fat: 14.2, satFat: 4.1, unsatFat: 9.2, transFat: 0,
    carbs: 2.5, fiber: 0.5, sugars: 1.0, netCarbs: 2.0, sodium: 320, potassium: 180, cholesterol: 350, calcium: 58, iron: 1.9
  },

  // --- VEGETABLES, CURRIES & PICKLES (Tarkari & Achar) ---
  {
    name: "Aloo Tama Bodi (Potato Bamboo Shoot Curry)",
    cal: 85, protein: 2.8, fat: 3.2, satFat: 0.5, unsatFat: 2.5, transFat: 0,
    carbs: 12.2, fiber: 2.8, sugars: 1.4, netCarbs: 9.4, sodium: 320, potassium: 280, cholesterol: 0, calcium: 28, iron: 1.5
  },
  {
    name: "Aloo Cauliflower Tarkari (Aloo Gobi)",
    cal: 92, protein: 2.2, fat: 4.1, satFat: 0.6, unsatFat: 3.3, transFat: 0,
    carbs: 12.5, fiber: 2.9, sugars: 2.1, netCarbs: 9.6, sodium: 290, potassium: 310, cholesterol: 0, calcium: 24, iron: 1.1
  },
  {
    name: "Saag (Rayoko Saag / Mustard Greens Stir Fry)",
    cal: 55, protein: 2.6, fat: 3.8, satFat: 0.5, unsatFat: 3.1, transFat: 0,
    carbs: 4.2, fiber: 2.5, sugars: 0.8, netCarbs: 1.7, sodium: 240, potassium: 320, cholesterol: 0, calcium: 115, iron: 2.1
  },
  {
    name: "Spinach Stir Fry (Palungo Saag)",
    cal: 48, protein: 2.8, fat: 3.2, satFat: 0.4, unsatFat: 2.6, transFat: 0,
    carbs: 3.6, fiber: 2.2, sugars: 0.4, netCarbs: 1.4, sodium: 220, potassium: 480, cholesterol: 0, calcium: 105, iron: 2.8
  },
  {
    name: "Aloo Dum (Spicy Potato)",
    cal: 125, protein: 2.1, fat: 5.5, satFat: 0.8, unsatFat: 4.4, transFat: 0,
    carbs: 17.5, fiber: 2.2, sugars: 1.2, netCarbs: 15.3, sodium: 360, potassium: 380, cholesterol: 0, calcium: 18, iron: 1.0
  },
  {
    name: "Aloo Kerau (Potato & Green Peas)",
    cal: 110, protein: 3.8, fat: 3.8, satFat: 0.6, unsatFat: 3.0, transFat: 0,
    carbs: 16.0, fiber: 3.8, sugars: 2.8, netCarbs: 12.2, sodium: 280, potassium: 320, cholesterol: 0, calcium: 22, iron: 1.3
  },
  {
    name: "Aloo Sesame Pickle (Aloo ko Achar with Til)",
    cal: 145, protein: 3.2, fat: 8.5, satFat: 1.2, unsatFat: 6.8, transFat: 0,
    carbs: 15.0, fiber: 3.1, sugars: 1.5, netCarbs: 11.9, sodium: 410, potassium: 340, cholesterol: 0, calcium: 85, iron: 2.2
  },
  {
    name: "Golbheda ko Achar (Tomato Chutney)",
    cal: 42, protein: 1.2, fat: 2.1, satFat: 0.3, unsatFat: 1.7, transFat: 0,
    carbs: 5.5, fiber: 1.4, sugars: 3.1, netCarbs: 4.1, sodium: 320, potassium: 210, cholesterol: 0, calcium: 15, iron: 0.8
  },
  {
    name: "Mula ko Achar (Fermented Radish Pickle)",
    cal: 35, protein: 1.0, fat: 1.8, satFat: 0.2, unsatFat: 1.5, transFat: 0,
    carbs: 4.2, fiber: 1.8, sugars: 1.5, netCarbs: 2.4, sodium: 680, potassium: 180, cholesterol: 0, calcium: 25, iron: 0.7
  },
  {
    name: "Lapsi ko Achar (Sweet & Sour Plum Pickle)",
    cal: 120, protein: 0.8, fat: 1.2, satFat: 0.2, unsatFat: 0.9, transFat: 0,
    carbs: 28.0, fiber: 2.2, sugars: 22.0, netCarbs: 25.8, sodium: 420, potassium: 140, cholesterol: 0, calcium: 30, iron: 1.2
  },
  {
    name: "Paneer Butter Masala",
    cal: 230, protein: 8.5, fat: 18.0, satFat: 9.2, unsatFat: 7.8, transFat: 0,
    carbs: 8.5, fiber: 1.2, sugars: 3.8, netCarbs: 7.3, sodium: 410, potassium: 180, cholesterol: 42, calcium: 210, iron: 0.9
  },
  {
    name: "Muttar Paneer",
    cal: 165, protein: 7.8, fat: 11.2, satFat: 5.5, unsatFat: 5.1, transFat: 0,
    carbs: 9.8, fiber: 2.5, sugars: 3.1, netCarbs: 7.3, sodium: 360, potassium: 220, cholesterol: 28, calcium: 180, iron: 1.2
  },
  {
    name: "Chana Masala",
    cal: 155, protein: 6.8, fat: 6.2, satFat: 0.9, unsatFat: 4.8, transFat: 0,
    carbs: 19.5, fiber: 5.2, sugars: 3.2, netCarbs: 14.3, sodium: 380, potassium: 290, cholesterol: 0, calcium: 45, iron: 2.4
  },
  {
    name: "Baigan Bharta (Eggplant Mash)",
    cal: 72, protein: 1.5, fat: 4.5, satFat: 0.6, unsatFat: 3.6, transFat: 0,
    carbs: 7.8, fiber: 3.1, sugars: 3.2, netCarbs: 4.7, sodium: 270, potassium: 230, cholesterol: 0, calcium: 22, iron: 0.7
  },
  {
    name: "Bhindi Fry (Lady Finger / Okra)",
    cal: 85, protein: 2.1, fat: 5.2, satFat: 0.7, unsatFat: 4.2, transFat: 0,
    carbs: 8.2, fiber: 3.2, sugars: 1.8, netCarbs: 5.0, sodium: 240, potassium: 260, cholesterol: 0, calcium: 75, iron: 0.9
  },
  {
    name: "Karela Fry (Bitter Gourd)",
    cal: 78, protein: 1.8, fat: 4.8, satFat: 0.6, unsatFat: 3.9, transFat: 0,
    carbs: 7.5, fiber: 2.8, sugars: 1.2, netCarbs: 4.7, sodium: 220, potassium: 290, cholesterol: 0, calcium: 28, iron: 1.0
  },
  {
    name: "Boiled Potato (Aloo)",
    cal: 87, protein: 1.9, fat: 0.1, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 20.1, fiber: 1.8, sugars: 0.9, netCarbs: 18.3, sodium: 6, potassium: 379, cholesterol: 0, calcium: 5, iron: 0.6
  },
  {
    name: "Sweet Potato Boiled (Sutkeri Aloo / Tarul)",
    cal: 76, protein: 1.4, fat: 0.2, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 17.7, fiber: 2.5, sugars: 5.7, netCarbs: 15.2, sodium: 27, potassium: 230, cholesterol: 0, calcium: 27, iron: 0.5
  },

  // --- DAIRY, CHEESE & CHHURPI ---
  {
    name: "Hard Chhurpi (Traditional Himalayan Cheese)",
    cal: 380, protein: 62.0, fat: 11.5, satFat: 7.2, unsatFat: 3.8, transFat: 0,
    carbs: 6.5, fiber: 0, sugars: 2.0, netCarbs: 6.5, sodium: 280, potassium: 180, cholesterol: 45, calcium: 820, iron: 1.8
  },
  {
    name: "Soft Chhurpi (Traditional)",
    cal: 165, protein: 18.5, fat: 7.2, satFat: 4.5, unsatFat: 2.3, transFat: 0,
    carbs: 4.2, fiber: 0, sugars: 3.1, netCarbs: 4.2, sodium: 120, potassium: 140, cholesterol: 25, calcium: 340, iron: 0.8
  },
  {
    name: "Paneer (Raw Indian Cottage Cheese)",
    cal: 265, protein: 18.3, fat: 20.8, satFat: 13.0, unsatFat: 6.8, transFat: 0,
    carbs: 1.2, fiber: 0, sugars: 1.2, netCarbs: 1.2, sodium: 18, potassium: 137, cholesterol: 60, calcium: 480, iron: 0.2
  },
  {
    name: "Whole Buffalo Milk",
    cal: 97, protein: 3.8, fat: 6.9, satFat: 4.6, unsatFat: 2.0, transFat: 0,
    carbs: 5.2, fiber: 0, sugars: 5.2, netCarbs: 5.2, sodium: 52, potassium: 178, cholesterol: 19, calcium: 169, iron: 0.1
  },
  {
    name: "Whole Cow Milk",
    cal: 61, protein: 3.2, fat: 3.3, satFat: 1.9, unsatFat: 1.2, transFat: 0,
    carbs: 4.8, fiber: 0, sugars: 4.8, netCarbs: 4.8, sodium: 43, potassium: 132, cholesterol: 10, calcium: 113, iron: 0.1
  },
  {
    name: "Dahi (Curd / Plain Yogurt)",
    cal: 61, protein: 3.5, fat: 3.3, satFat: 2.1, unsatFat: 1.0, transFat: 0,
    carbs: 4.7, fiber: 0, sugars: 4.7, netCarbs: 4.7, sodium: 46, potassium: 155, cholesterol: 13, calcium: 121, iron: 0.1
  },
  {
    name: "Sweet Lassi",
    cal: 110, protein: 3.1, fat: 3.0, satFat: 1.8, unsatFat: 1.0, transFat: 0,
    carbs: 17.5, fiber: 0, sugars: 16.5, netCarbs: 17.5, sodium: 40, potassium: 140, cholesterol: 10, calcium: 105, iron: 0.1
  },
  {
    name: "Mahi / Mohi (Buttermilk)",
    cal: 40, protein: 3.3, fat: 0.9, satFat: 0.5, unsatFat: 0.3, transFat: 0,
    carbs: 4.8, fiber: 0, sugars: 4.8, netCarbs: 4.8, sodium: 105, potassium: 151, cholesterol: 4, calcium: 116, iron: 0.1
  },
  {
    name: "Ghee (Clarified Butter - 1 tbsp ~13g / values for 100g)",
    cal: 884, protein: 0.3, fat: 99.5, satFat: 61.9, unsatFat: 32.5, transFat: 0,
    carbs: 0, fiber: 0, sugars: 0, netCarbs: 0, sodium: 2, potassium: 5, cholesterol: 256, calcium: 6, iron: 0
  },

  // --- SWEETS & DESSERTS (Mithai & Khaja) ---
  {
    name: "Gulab Jamun (2 pcs ~100g)",
    cal: 300, protein: 4.2, fat: 10.5, satFat: 5.2, unsatFat: 4.5, transFat: 0,
    carbs: 48.0, fiber: 0.5, sugars: 38.0, netCarbs: 47.5, sodium: 85, potassium: 90, cholesterol: 18, calcium: 95, iron: 0.5
  },
  {
    name: "Rasgulla (2 pcs ~100g)",
    cal: 186, protein: 4.0, fat: 1.8, satFat: 1.1, unsatFat: 0.6, transFat: 0,
    carbs: 38.5, fiber: 0, sugars: 33.0, netCarbs: 38.5, sodium: 25, potassium: 45, cholesterol: 6, calcium: 120, iron: 0.3
  },
  {
    name: "Jalebi",
    cal: 310, protein: 2.1, fat: 9.2, satFat: 2.8, unsatFat: 5.8, transFat: 0,
    carbs: 56.0, fiber: 0.8, sugars: 42.0, netCarbs: 55.2, sodium: 60, potassium: 40, cholesterol: 0, calcium: 20, iron: 0.8
  },
  {
    name: "Laddoo (Motichoor)",
    cal: 385, protein: 5.5, fat: 16.2, satFat: 6.5, unsatFat: 8.8, transFat: 0,
    carbs: 54.0, fiber: 2.1, sugars: 39.0, netCarbs: 51.9, sodium: 45, potassium: 110, cholesterol: 12, calcium: 35, iron: 1.4
  },
  {
    name: "Kheer (Rice Pudding)",
    cal: 142, protein: 3.8, fat: 4.2, satFat: 2.5, unsatFat: 1.4, transFat: 0,
    carbs: 22.5, fiber: 0.3, sugars: 15.2, netCarbs: 22.2, sodium: 55, potassium: 160, cholesterol: 14, calcium: 125, iron: 0.3
  },
  {
    name: "Halwa (Suji / Semolina)",
    cal: 285, protein: 4.5, fat: 11.2, satFat: 6.2, unsatFat: 4.3, transFat: 0,
    carbs: 42.0, fiber: 1.2, sugars: 24.0, netCarbs: 40.8, sodium: 35, potassium: 85, cholesterol: 20, calcium: 30, iron: 1.1
  },
  {
    name: "Suji Puri (Sweet Sheera)",
    cal: 270, protein: 4.0, fat: 10.0, satFat: 5.5, unsatFat: 3.8, transFat: 0,
    carbs: 41.0, fiber: 1.0, sugars: 22.0, netCarbs: 40.0, sodium: 30, potassium: 75, cholesterol: 18, calcium: 28, iron: 0.9
  },

  // --- NUTS, SEEDS & OILS ---
  {
    name: "Peanuts / Badam (Roasted Salted)",
    cal: 585, protein: 23.7, fat: 49.7, satFat: 6.9, unsatFat: 40.2, transFat: 0,
    carbs: 21.5, fiber: 8.0, sugars: 4.2, netCarbs: 13.5, sodium: 410, potassium: 658, cholesterol: 0, calcium: 54, iron: 2.3
  },
  {
    name: "Almonds (Badaam)",
    cal: 579, protein: 21.2, fat: 49.9, satFat: 3.8, unsatFat: 42.1, transFat: 0,
    carbs: 21.6, fiber: 12.5, sugars: 4.4, netCarbs: 9.1, sodium: 1, potassium: 733, cholesterol: 0, calcium: 269, iron: 3.7
  },
  {
    name: "Walnuts (Okhar)",
    cal: 654, protein: 15.2, fat: 65.2, satFat: 6.1, unsatFat: 56.3, transFat: 0,
    carbs: 13.7, fiber: 6.7, sugars: 2.6, netCarbs: 7.0, sodium: 2, potassium: 441, cholesterol: 0, calcium: 98, iron: 2.9
  },
  {
    name: "Cashews (Kaju)",
    cal: 553, protein: 18.2, fat: 43.8, satFat: 7.8, unsatFat: 33.8, transFat: 0,
    carbs: 30.2, fiber: 3.3, sugars: 5.9, netCarbs: 26.9, sodium: 12, potassium: 660, cholesterol: 0, calcium: 37, iron: 6.7
  },
  {
    name: "Sesame Seeds (Til)",
    cal: 573, protein: 17.7, fat: 49.7, satFat: 7.0, unsatFat: 40.5, transFat: 0,
    carbs: 23.4, fiber: 11.8, sugars: 0.3, netCarbs: 11.6, sodium: 11, potassium: 468, cholesterol: 0, calcium: 975, iron: 14.6
  },
  {
    name: "Mustard Oil (Tori ko Tel)",
    cal: 884, protein: 0, fat: 100, satFat: 11.6, unsatFat: 83.2, transFat: 0,
    carbs: 0, fiber: 0, sugars: 0, netCarbs: 0, sodium: 0, potassium: 0, cholesterol: 0, calcium: 0, iron: 0
  },
  {
    name: "Sunflower Oil",
    cal: 884, protein: 0, fat: 100, satFat: 10.3, unsatFat: 84.7, transFat: 0,
    carbs: 0, fiber: 0, sugars: 0, netCarbs: 0, sodium: 0, potassium: 0, cholesterol: 0, calcium: 0, iron: 0
  },

  // --- LOCAL FRUITS & FRESH PRODUCE ---
  {
    name: "Banana (Kera)",
    cal: 89, protein: 1.1, fat: 0.3, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 22.8, fiber: 2.6, sugars: 12.2, netCarbs: 20.2, sodium: 1, potassium: 358, cholesterol: 0, calcium: 5, iron: 0.3
  },
  {
    name: "Apple (Aabu / Syau)",
    cal: 52, protein: 0.3, fat: 0.2, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 13.8, fiber: 2.4, sugars: 10.4, netCarbs: 11.4, sodium: 1, potassium: 107, cholesterol: 0, calcium: 6, iron: 0.1
  },
  {
    name: "Mango (Aam)",
    cal: 60, protein: 0.8, fat: 0.4, satFat: 0.1, unsatFat: 0.2, transFat: 0,
    carbs: 15.0, fiber: 1.6, sugars: 13.7, netCarbs: 13.4, sodium: 1, potassium: 168, cholesterol: 0, calcium: 11, iron: 0.2
  },
  {
    name: "Guava (Amba)",
    cal: 68, protein: 2.6, fat: 1.0, satFat: 0.3, unsatFat: 0.5, transFat: 0,
    carbs: 14.3, fiber: 5.4, sugars: 8.9, netCarbs: 8.9, sodium: 2, potassium: 417, cholesterol: 0, calcium: 18, iron: 0.3
  },
  {
    name: "Papaya (Mewa)",
    cal: 43, protein: 0.5, fat: 0.3, satFat: 0.1, unsatFat: 0.1, transFat: 0,
    carbs: 10.8, fiber: 1.7, sugars: 7.8, netCarbs: 9.1, sodium: 8, potassium: 182, cholesterol: 0, calcium: 20, iron: 0.25
  },
  {
    name: "Orange (Suntala)",
    cal: 47, protein: 0.9, fat: 0.1, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 11.8, fiber: 2.4, sugars: 9.4, netCarbs: 9.4, sodium: 0, potassium: 181, cholesterol: 0, calcium: 40, iron: 0.1
  },
  {
    name: "Cucumber (Kankro)",
    cal: 15, protein: 0.7, fat: 0.1, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 3.6, fiber: 0.5, sugars: 1.7, netCarbs: 3.1, sodium: 2, potassium: 147, cholesterol: 0, calcium: 16, iron: 0.3
  },
  {
    name: "Radish (Mula - Raw)",
    cal: 16, protein: 0.7, fat: 0.1, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 3.4, fiber: 1.6, sugars: 1.9, netCarbs: 1.8, sodium: 39, potassium: 233, cholesterol: 0, calcium: 25, iron: 0.3
  },

  // --- BEVERAGES & TEA ---
  {
    name: "Nepalese Milk Tea with Sugar (Chiya)",
    cal: 75, protein: 2.1, fat: 2.2, satFat: 1.3, unsatFat: 0.8, transFat: 0,
    carbs: 11.8, fiber: 0, sugars: 11.2, netCarbs: 11.8, sodium: 32, potassium: 95, cholesterol: 7, calcium: 75, iron: 0.1
  },
  {
    name: "Black Tea (Kalo Chiya - Plain)",
    cal: 2, protein: 0, fat: 0, satFat: 0, unsatFat: 0, transFat: 0,
    carbs: 0.4, fiber: 0, sugars: 0, netCarbs: 0.4, sodium: 3, potassium: 37, cholesterol: 0, calcium: 0, iron: 0.02
  },
  {
    name: "Butter Tea (Suja Chiya - Tibetan/Sherpa)",
    cal: 115, protein: 1.2, fat: 11.5, satFat: 7.2, unsatFat: 3.8, transFat: 0,
    carbs: 1.8, fiber: 0, sugars: 1.0, netCarbs: 1.8, sodium: 180, potassium: 50, cholesterol: 30, calcium: 35, iron: 0.1
  },
  {
    name: "Instant Coffee with Milk & Sugar",
    cal: 85, protein: 2.2, fat: 2.1, satFat: 1.3, unsatFat: 0.7, transFat: 0,
    carbs: 14.0, fiber: 0, sugars: 13.1, netCarbs: 14.0, sodium: 35, potassium: 110, cholesterol: 7, calcium: 78, iron: 0.1
  },
  {
    name: "Chicken Breast (Cooked, Skinless)",
    cal: 165, protein: 31.0, fat: 3.6, satFat: 1.0, unsatFat: 2.3, transFat: 0,
    carbs: 0.0, fiber: 0, sugars: 0.0, netCarbs: 0.0, sodium: 74, potassium: 256, cholesterol: 85, calcium: 15, iron: 1.0
  },
  {
    name: "Whole Egg (Cooked, Hard-Boiled)",
    cal: 155, protein: 12.6, fat: 10.6, satFat: 3.3, unsatFat: 6.3, transFat: 0,
    carbs: 1.1, fiber: 0, sugars: 1.1, netCarbs: 1.1, sodium: 124, potassium: 126, cholesterol: 373, calcium: 50, iron: 1.2
  },
  {
    name: "Egg Whites (Cooked)",
    cal: 52, protein: 10.9, fat: 0.2, satFat: 0.0, unsatFat: 0.1, transFat: 0,
    carbs: 0.7, fiber: 0, sugars: 0.7, netCarbs: 0.7, sodium: 166, potassium: 163, cholesterol: 0, calcium: 7, iron: 0.1
  },
  {
    name: "Greek Yogurt (Plain, Non-Fat)",
    cal: 59, protein: 10.0, fat: 0.4, satFat: 0.1, unsatFat: 0.2, transFat: 0,
    carbs: 3.6, fiber: 0, sugars: 3.2, netCarbs: 3.6, sodium: 36, potassium: 141, cholesterol: 5, calcium: 110, iron: 0.1
  },
  {
    name: "Cottage Cheese (Low Fat 2%)",
    cal: 82, protein: 11.0, fat: 2.3, satFat: 1.4, unsatFat: 0.8, transFat: 0,
    carbs: 3.4, fiber: 0, sugars: 2.7, netCarbs: 3.4, sodium: 330, potassium: 104, cholesterol: 9, calcium: 83, iron: 0.1
  },
  {
    name: "Tuna (Canned in Water, Drained)",
    cal: 116, protein: 25.5, fat: 1.0, satFat: 0.2, unsatFat: 0.6, transFat: 0,
    carbs: 0.0, fiber: 0, sugars: 0.0, netCarbs: 0.0, sodium: 338, potassium: 237, cholesterol: 30, calcium: 11, iron: 1.0
  },
  {
    name: "Salmon (Cooked, Baked)",
    cal: 206, protein: 22.0, fat: 12.3, satFat: 2.5, unsatFat: 8.6, transFat: 0,
    carbs: 0.0, fiber: 0, sugars: 0.0, netCarbs: 0.0, sodium: 61, potassium: 384, cholesterol: 63, calcium: 12, iron: 0.8
  },
  {
    name: "Firm Tofu",
    cal: 83, protein: 10.0, fat: 5.3, satFat: 0.8, unsatFat: 4.1, transFat: 0,
    carbs: 1.9, fiber: 0.9, sugars: 0.5, netCarbs: 1.0, sodium: 14, potassium: 121, cholesterol: 0, calcium: 282, iron: 2.7
  },
  {
    name: "Lentils (Cooked, Boiled)",
    cal: 116, protein: 9.0, fat: 0.4, satFat: 0.1, unsatFat: 0.2, transFat: 0,
    carbs: 20.0, fiber: 7.9, sugars: 1.8, netCarbs: 12.1, sodium: 2, potassium: 369, cholesterol: 0, calcium: 19, iron: 3.3
  },
  {
    name: "Whey Protein Isolate Powder",
    cal: 370, protein: 80.0, fat: 2.0, satFat: 1.0, unsatFat: 0.8, transFat: 0,
    carbs: 3.0, fiber: 0, sugars: 1.5, netCarbs: 3.0, sodium: 170, potassium: 420, cholesterol: 10, calcium: 450, iron: 0.5
  },
  {
    name: "Peanut Butter (Smooth / Whole Nut)",
    cal: 588, protein: 25.0, fat: 50.0, satFat: 10.0, unsatFat: 40.0, transFat: 0,
    carbs: 20.0, fiber: 6.0, sugars: 9.0, netCarbs: 14.0, sodium: 17, potassium: 649, cholesterol: 0, calcium: 43, iron: 1.9
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
  if (calculatorForm) calculatorForm.addEventListener('submit', handleTargetCalculation);
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

// --- Goal & Macro Calculator (Mifflin-St Jeor Formula) ---
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
  const targetFiber = Math.round((targetCal / 1000) * 14);

  dailyTargets = {
    cal: targetCal,
    protein: targetProtein,
    carbs: targetCarbs,
    fat: targetFat,
    fiber: targetFiber
  };

  // Update UI & LocalStorage
  const badge = document.getElementById('target-mode-badge');
  if (badge) badge.textContent = `Personal Goal (${goal.toUpperCase()})`;
  saveToLocalStorage();
  updateDashboard();
  
  alert(`Goals Updated Successfully!\nDaily Calories: ${targetCal} kcal\nProtein: ${targetProtein}g | Carbs: ${targetCarbs}g | Fat: ${targetFat}g`);
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
      <div class="suggestion-item" onclick="selectSuggestion('${item.name.replace(/'/g, "\\'")}')">${item.name}</div>
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
    // Fallback item for unlisted foods
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

// --- Clean, Excel-Friendly CSV Export ---
function exportCleanCSV() {
  if (loggedItems.length === 0) {
    alert("No logged food items to export!");
    return;
  }

  const currentDate = new Date().toLocaleDateString();
  
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

  let csvRows = [
    `"NUTRIPULSE DAILY NUTRITION LOG"`,
    `"Date Exported:","${currentDate}"`,
    `"Daily Calorie Target:","${dailyTargets.cal} kcal"`,
    `"Daily Protein Target:","${dailyTargets.protein} g"`,
    `""`,
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

  csvRows.push(`""`);
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
      const badge = document.getElementById('target-mode-badge');
      if (badge) badge.textContent = 'Custom Goals Loaded';
    } catch (e) {}
  }
}
// Combiner State Variable
let selectedCombinerIds = new Set();

// Add inside setupEventListeners()
document.getElementById('clear-combiner-btn')?.addEventListener('click', clearCombinerSelection);

// Function to Render Item List in Combiner Card
function renderCombinerItems() {
  const container = document.getElementById('combiner-items-list');
  if (!container) return;

  if (loggedItems.length === 0) {
    container.innerHTML = `<p style="color: var(--text-muted); font-size: 0.85rem; text-align: center; padding: 12px;">Log food items above to combine them.</p>`;
    calculateCombinerTotals();
    return;
  }

  container.innerHTML = loggedItems.map(item => `
    <label class="combiner-item-row">
      <input type="checkbox" 
             value="${item.id}" 
             ${selectedCombinerIds.has(item.id) ? 'checked' : ''} 
             onchange="toggleCombinerItem(${item.id})">
      <div class="combiner-item-info">
        <span><strong>${item.name}</strong> (${item.weightEntered}${item.unit})</span>
        <span>${item.cal} kcal</span>
      </div>
    </label>
  `).join('');

  calculateCombinerTotals();
}

// Toggle selection state
function toggleCombinerItem(id) {
  if (selectedCombinerIds.has(id)) {
    selectedCombinerIds.delete(id);
  } else {
    selectedCombinerIds.add(id);
  }
  calculateCombinerTotals();
}

// Calculate combined totals
function calculateCombinerTotals() {
  const selectedItems = loggedItems.filter(item => selectedCombinerIds.has(item.id));

  const totals = selectedItems.reduce((acc, curr) => ({
    cal: acc.cal + curr.cal,
    protein: acc.protein + curr.protein,
    carbs: acc.carbs + curr.carbs,
    fat: acc.fat + curr.fat,
    fiber: acc.fiber + curr.fiber
  }), { cal: 0, protein: 0, carbs: 0, fat: 0, fiber: 0 });

  document.getElementById('combine-cal').innerHTML = `${totals.cal.toLocaleString()} <small>kcal</small>`;
  document.getElementById('combine-protein').innerHTML = `${totals.protein.toFixed(1)} <small>g</small>`;
  document.getElementById('combine-carbs').innerHTML = `${totals.carbs.toFixed(1)} <small>g</small>`;
  document.getElementById('combine-fat').innerHTML = `${totals.fat.toFixed(1)} <small>g</small>`;
  document.getElementById('combine-fiber').innerHTML = `${totals.fiber.toFixed(1)} <small>g</small>`;
}

// Reset Combiner Selection
function clearCombinerSelection() {
  selectedCombinerIds.clear();
  renderCombinerItems();
}

// Ensure renderCombinerItems() is called inside your existing rendering/logging functions:
// 1. Inside DOMContentLoaded
// 2. Inside logItem()
// 3. Inside deleteItem()
// 4. Inside clearAllLog()