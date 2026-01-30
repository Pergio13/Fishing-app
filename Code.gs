/**
 * FISHING FORECAST WEB APP v4.0
 * Με πλήρη λειτουργίες, ιστορικό, αγαπημένα και καλύτερο UI
 */

// ==================== MAIN WEB APP ====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('🎣 Fishing Forecast Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== LOCATIONS API ====================
function getLocations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Locations");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    data.shift(); // Remove header
    return data.map(row => ({
      name: row[0] || "",
      lat: parseFloat(row[1]) || 0,
      lon: parseFloat(row[2]) || 0,
      winMin: parseFloat(row[3]) || 0,
      winMax: parseFloat(row[4]) || 360,
      favorite: row[5] === true || row[5] === "TRUE" || false
    }));
  } catch (error) {
    console.error("Error in getLocations:", error);
    return [];
  }
}

function getFavoriteLocations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Locations");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    data.shift();
    return data
      .filter(row => row[5] === true || row[5] === "TRUE")
      .map(row => ({
        name: row[0] || "",
        lat: parseFloat(row[1]) || 0,
        lon: parseFloat(row[2]) || 0,
        winMin: parseFloat(row[3]) || 0,
        winMax: parseFloat(row[4]) || 360,
        favorite: true
      }));
  } catch (error) {
    console.error("Error in getFavoriteLocations:", error);
    return [];
  }
}

function toggleFavorite(locationName, isFavorite) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Locations");
    if (!sheet) return { success: false, error: "Sheet not found" };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === locationName) {
        sheet.getRange(i + 1, 6).setValue(isFavorite);
        return { success: true };
      }
    }
    return { success: false, error: "Location not found" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ==================== CUSTOM PINS API ====================
function saveCustomPin(pinData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CustomPins");
    
    if (!sheet) {
      sheet = ss.insertSheet("CustomPins");
      sheet.appendRow(["Name", "Lat", "Lon", "Timestamp", "Color", "Type"]);
    }
    
    sheet.appendRow([
      pinData.name || "Custom Pin",
      pinData.lat,
      pinData.lon,
      new Date(),
      pinData.color || "#FF5722",
      pinData.type || "custom" // "custom" or "fish"
    ]);
    
    return { success: true, id: sheet.getLastRow() - 1 };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
function getCustomPins() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CustomPins");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    data.shift();
    return data.map((row, index) => ({
      id: index,
      name: row[0] || "Custom Pin",
      lat: parseFloat(row[1]) || 0,
      lon: parseFloat(row[2]) || 0,
      timestamp: row[3] || new Date(),
      color: row[4] || "#FF5722",
      type: row[5] || "custom"
    }));
  } catch (error) {
    console.error("Error in getCustomPins:", error);
    return [];
  }
}

function deleteCustomPin(pinId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CustomPins");
    if (!sheet) return { success: false, error: "Sheet not found" };
    
    // +2 because header row and 0-index offset
    sheet.deleteRow(pinId + 2);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ==================== FORECAST API ====================
function getForecastForCoordinates(lat, lon) {
  try {
    const location = {
      name: `📍 ${lat.toFixed(4)}, ${lon.toFixed(4)}`,
      lat: lat,
      lon: lon,
      winMin: 0,
      winMax: 360,
      isValid: true
    };
    
    const weatherData = fetchWeatherData(location);
    if (!weatherData) {
      throw new Error("Failed to fetch weather data");
    }
    
    const forecast = createForecastForWeb(weatherData, location);
    
    return {
      success: true,
      location: location.name,
      coordinates: { lat, lon },
      forecast: forecast,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error("Error in getForecastForCoordinates:", error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getForecastForLocation(locationName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const forecastSheet = ss.getSheetByName("Πρόγνωση");
    const locSheet = ss.getSheetByName("Locations");
    
    if (!forecastSheet || !locSheet) {
      throw new Error("Required sheets not found");
    }
    
    // Save selected location
    forecastSheet.getRange("B1").setValue(locationName);
    
    // Get location data
    const location = getLocationDataByName(locationName, locSheet);
    if (!location.isValid) {
      throw new Error("Invalid location data");
    }
    
    const weatherData = fetchWeatherData(location);
    if (!weatherData) {
      throw new Error("Failed to fetch weather data");
    }
    
    const forecast = createForecastForWeb(weatherData, location);
    
    // Also update the sheet (optional)
    const results = createForecastMatrix(weatherData, location);
    displayForecastResults(forecastSheet, results);
    
    return {
      success: true,
      location: location.name,
      coordinates: { lat: location.lat, lon: location.lon },
      forecast: forecast,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error("Error in getForecastForLocation:", error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getLocationDataByName(locationName, locSheet) {
  const locData = locSheet.getDataRange().getValues();
  for (let i = 1; i < locData.length; i++) {
    if (locData[i][0] && locData[i][0].toString().trim() === locationName) {
      return {
        name: locData[i][0],
        lat: parseFloat(locData[i][1]) || 0,
        lon: parseFloat(locData[i][2]) || 0,
        winMin: parseFloat(locData[i][3]) || 0,
        winMax: parseFloat(locData[i][4]) || 360,
        isValid: true
      };
    }
  }
  return { isValid: false };
}

function fetchWeatherData(loc) {
  try {
    const weatherUrl = `https://api.open-meteo.com/v1/forecast?latitude=${loc.lat}&longitude=${loc.lon}&hourly=surface_pressure,temperature_2m,wind_speed_10m,wind_direction_10m,precipitation,cloud_cover&past_days=1&forecast_days=7&timezone=auto`;
    
    const marineUrl = `https://marine-api.open-meteo.com/v1/marine?latitude=${loc.lat}&longitude=${loc.lon}&hourly=wave_height,sea_surface_temperature`;
    
    const weatherResponse = UrlFetchApp.fetch(weatherUrl, { muteHttpExceptions: true });
    const marineResponse = UrlFetchApp.fetch(marineUrl, { muteHttpExceptions: true });
    
    if (weatherResponse.getResponseCode() !== 200) {
      throw new Error(`Weather API error: ${weatherResponse.getResponseCode()}`);
    }
    
    const weatherData = JSON.parse(weatherResponse.getContentText()).hourly;
    let marineData = { wave_height: [], sea_surface_temperature: [] };
    
    if (marineResponse.getResponseCode() === 200) {
      marineData = JSON.parse(marineResponse.getContentText()).hourly;
    }
    
    return {
      weather: weatherData,
      marine: marineData,
      location: loc
    };
    
  } catch (error) {
    console.error("Error in fetchWeatherData:", error);
    return null;
  }
}
function createForecastForWeb(data, location) {
  const { weather, marine } = data;
  const greekDays = ["Κυρ", "Δευ", "Τρι", "Τετ", "Πέμ", "Παρ", "Σάβ"];
  const timeSlots = ["06:00", "09:00", "12:00", "15:00", "18:00", "21:00"];
  
  const forecast = {
    location: location.name,
    coordinates: { lat: location.lat, lon: location.lon },
    days: [],
    summary: {
      bestDay: null,
      bestScore: 0,
      bestTime: "",
      hasRain: false
    }
  };
  
  // Process each day
  for (let dayIndex = 1; dayIndex <= 7; dayIndex++) {
    const hourIndex = dayIndex * 24;
    const date = new Date(weather.time[hourIndex]);
    const moonIllum = getMoonIllumination(date);
    
    const dayData = {
      date: date.toISOString().split('T')[0],
      dayName: greekDays[date.getDay()],
      formattedDate: `${greekDays[date.getDay()]} ${date.getDate()}/${date.getMonth() + 1}`,
      moonIllumination: moonIllum,
      moonEmoji: getMoonPhaseEmoji(moonIllum),
      timeSlots: [],
      dayScore: 0,
      bestTimeSlot: null,
      isRaining: false,
      dayColor: "#FFFFFF"
    };
    
    let dayTotalScore = 0;
    let daySlotCount = 0;
    
    // Process each time slot
    for (let slotIndex = 0; slotIndex < timeSlots.length; slotIndex++) {
      const hour = parseInt(timeSlots[slotIndex].split(':')[0]);
      const dataIndex = hourIndex + hour;
      
      const dataPoint = {
        time: timeSlots[slotIndex],
        timestamp: weather.time[dataIndex],
        airTemp: weather.temperature_2m[dataIndex] || 0,
        seaTemp: marine.sea_surface_temperature?.[dataIndex] || 0,
        pressure: weather.surface_pressure[dataIndex] || 1013,
        pressureTrend: weather.surface_pressure[dataIndex] && weather.surface_pressure[dataIndex - 12] 
          ? weather.surface_pressure[dataIndex] - weather.surface_pressure[dataIndex - 12] 
          : 0,
        windSpeed: weather.wind_speed_10m[dataIndex] || 0,
        windDirection: weather.wind_direction_10m[dataIndex] || 0,
        waveHeight: marine.wave_height?.[dataIndex] || 0,
        rain: weather.precipitation[dataIndex] || 0,
        cloudCover: weather.cloud_cover[dataIndex] || 0
      };
      
      // Calculate score
      const scoreBreakdown = calculateVidalisScore(dataPoint, location, moonIllum);
      const finalScore = Math.min(scoreBreakdown.total / 100, 1);
      const scorePercent = Math.round(finalScore * 100);
      
      const timeSlot = {
        ...dataPoint,
        score: finalScore,
        scorePercent: scorePercent,
        scoreColor: getScoreColor(finalScore, dataPoint.rain),
        scoreEmoji: getScoreEmoji(finalScore, dataPoint.rain),
        reasons: scoreBreakdown.reasons
      };
      
      dayData.timeSlots.push(timeSlot);
      dayTotalScore += scorePercent;
      daySlotCount++;
      
      // Track best time slot for the day
      if (!dayData.bestTimeSlot || scorePercent > dayData.bestTimeSlot.scorePercent) {
        dayData.bestTimeSlot = timeSlot;
      }
      
      // Check for rain
      if (dataPoint.rain > 0.4) dayData.isRaining = true;
    }
    
    // Calculate average day score
    dayData.dayScore = Math.round(dayTotalScore / daySlotCount);
    dayData.dayColor = getScoreColor(dayData.dayScore / 100, 0);
    
    // Update summary if this is the best day
    if (dayData.dayScore > forecast.summary.bestScore) {
      forecast.summary.bestDay = dayData.formattedDate;
      forecast.summary.bestScore = dayData.dayScore;
      forecast.summary.bestTime = dayData.bestTimeSlot.time;
    }
    
    if (dayData.isRaining) forecast.summary.hasRain = true;
    
    forecast.days.push(dayData);
  }
  
  return forecast;
}

// ==================== CATCH LOGGING API ====================
function logCatch(catchData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let catchSheet = ss.getSheetByName("CatchLog");
    
    if (!catchSheet) {
      // Create sheet if it doesn't exist
      catchSheet = ss.insertSheet("CatchLog");
      catchSheet.appendRow([
        "Timestamp", "Location", "Latitude", "Longitude", "Fish Name", "Weight", 
        "Score", "Pressure", "Pressure Trend", "Wind Speed", "Wave Height", 
        "Moon Illumination", "Air Temp", "Sea Temp", "Notes", "Weather Conditions"
      ]);
    }
    
    // Create new row
    const newRow = [
      new Date(),
      catchData.locationName || "Unknown Location",
      catchData.latitude || 0,
      catchData.longitude || 0,
      catchData.fishName,
      parseFloat(catchData.weight) || 0,
      catchData.score || 0,
      catchData.pressure || 0,
      catchData.pressureTrend || 0,
      catchData.windSpeed || 0,
      catchData.waveHeight || 0,
      catchData.moonIllumination || 0,
      catchData.airTemp || 0,
      catchData.seaTemp || 0,
      catchData.notes || "",
      catchData.weatherConditions || ""
    ];
    
    catchSheet.appendRow(newRow);
    
    // Format the new row
    const lastRow = catchSheet.getLastRow();
    formatCatchRow(catchSheet, lastRow);
    
    return {
      success: true,
      message: `🎣 Καταγράφηκε ${catchData.fishName} (${catchData.weight}kg)`,
      rowNumber: lastRow
    };
    
  } catch (error) {
    console.error("Error in logCatch:", error);
    return {
      success: false,
      error: error.message
    };
  }
}

function getCatchHistory(limit = 50) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catchSheet = ss.getSheetByName("CatchLog");
    
    if (!catchSheet) {
      return [];
    }
    
    const data = catchSheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const rows = data.slice(1).reverse().slice(0, limit);
    
    return rows.map(row => {
      const entry = {};
      headers.forEach((header, index) => {
        entry[header] = row[index];
      });
      return entry;
    });
    
  } catch (error) {
    console.error("Error in getCatchHistory:", error);
    return [];
  }
}

function formatCatchRow(sheet, rowNumber) {
  const rowRange = sheet.getRange(rowNumber, 1, 1, 16);
  
  // Format date
  sheet.getRange(rowNumber, 1).setNumberFormat("dd/MM/yyyy HH:mm");
  
  // Format weight
  sheet.getRange(rowNumber, 6).setNumberFormat("0.00");
  
  // Format percentages
  sheet.getRange(rowNumber, 7).setNumberFormat("0%"); // Score
  sheet.getRange(rowNumber, 12).setNumberFormat("0%"); // Moon
  
  // Color based on weight
  const weight = sheet.getRange(rowNumber, 6).getValue();
  let bgColor = "#FFFFFF";
  
  if (weight >= 5) bgColor = "#D4EDDA"; // Green for big fish
  else if (weight >= 2) bgColor = "#FFF3CD"; // Yellow for medium fish
  
  rowRange.setBackground(bgColor);
  rowRange.setHorizontalAlignment("center");
  rowRange.setVerticalAlignment("middle");
  
  // Add border
  rowRange.setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
}

// ==================== HELPER FUNCTIONS ====================
function calculateVidalisScore(dataPoint, location, moonIllum) {
  let score = 0;
  let reasons = [];
  
  // Pressure drop
  if (dataPoint.pressureTrend < -1.5) {
    score += 35;
    reasons.push("📉 Πτώση πίεσης");
  }
  
  // Low pressure
  if (dataPoint.pressure < 1010) {
    score += 20;
    reasons.push("🌡️ Χαμηλή πίεση");
  }
  
  // Moon phase
  if (moonIllum <= 0.15 || moonIllum >= 0.85) {
    score += 30;
    reasons.push("🌕 Ισχυρή σελήνη");
  }
  
  // Optimal wind and waves
  if (dataPoint.windDirection >= location.winMin && 
      dataPoint.windDirection <= location.winMax && 
      dataPoint.waveHeight > 0.5) {
    score += 15;
    reasons.push("🌊 Καλό κύμα");
  }
  
  // Cloud penalty
  if (dataPoint.cloudCover > 80) {
    score -= 10;
  }
  
  // Rain penalty
  if (dataPoint.rain > 1.5) {
    score -= 20;
    reasons.push("🌧️ Ισχυρή βροχή");
  }
  
  return {
    total: Math.max(0, score),
    reasons: reasons
  };
}

function getMoonIllumination(date) {
  const knownNewMoon = new Date(Date.UTC(2024, 0, 11, 11, 57, 0));
  const targetDate = new Date(date.getTime() + date.getTimezoneOffset() * 60000);
  
  const diffMs = targetDate - knownNewMoon;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  const moonCycle = 29.53058867;
  
  const phase = ((diffDays % moonCycle) + moonCycle) % moonCycle;
  const illumination = 0.5 * (1 - Math.cos(2 * Math.PI * phase / moonCycle));
  
  return Math.max(0, Math.min(1, illumination));
}

function getMoonPhaseEmoji(illumination) {
  if (illumination < 0.1) return "🌑";
  if (illumination < 0.4) return "🌘";
  if (illumination < 0.6) return "🌗";
  if (illumination < 0.9) return "🌖";
  return "🌕";
}

function getScoreEmoji(score, rain) {
  if (rain > 1.5) return "🌧️";
  if (score >= 0.7) return "🎯";
  if (score >= 0.45) return "✅";
  return "❌";
}

function getScoreColor(score, rain) {
  // Χωρίς μπλε για βροχή - μόνο το εικονίδιο δείχνει βροχή
  if (score >= 0.7) return "#D4EDDA"; // Green for excellent
  if (score >= 0.45) return "#FFF3CD"; // Yellow for good
  return "#F8D7DA"; // Red for poor
}

// ==================== LEGACY FUNCTIONS (for Google Sheets) ====================
function getFishingForecast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const forecastSheet = ss.getSheetByName("Πρόγνωση");
  const locSheet = ss.getSheetByName("Locations");
  
  prepareForecastSheet(forecastSheet);
  const location = getLocationData(forecastSheet, locSheet);
  if (!location.isValid) return;

  const weatherData = fetchWeatherData(location);
  if (!weatherData) return;

  const results = createForecastMatrix(weatherData, location);
  displayForecastResults(forecastSheet, results);
}

function getLocationData(forecastSheet, locSheet) {
  const selectedLoc = forecastSheet.getRange("B1").getValue().toString().trim();
  const locData = locSheet.getDataRange().getValues();
  for (let i = 1; i < locData.length; i++) {
    if (locData[i][0].toString().trim() === selectedLoc) {
      return { 
        name: locData[i][0], 
        lat: locData[i][1], 
        lon: locData[i][2], 
        winMin: locData[i][3], 
        winMax: locData[i][4], 
        isValid: true 
      };
    }
  }
  return { isValid: false };
}

function createForecastMatrix(data, loc) {
  const { weather, marine } = data;
  const timeSlots = ["06:00", "09:00", "12:00", "15:00", "18:00", "21:00"];
  let matrix = timeSlots.map(s => [s]);
  let colors = timeSlots.map(() => [null]);
  
  for (let d = 1; d <= 7; d++) {
    const dayIdx = d * 24;
    const moonIllum = getMoonIllumination(new Date(weather.time[dayIdx]));
    
    for (let t = 0; t < timeSlots.length; t++) {
      const idx = dayIdx + parseInt(timeSlots[t].split(':')[0]);
      const dp = {
        pressure: weather.surface_pressure[idx] || 1013,
        pressureTrend: (weather.surface_pressure[idx] || 1013) - (weather.surface_pressure[idx-12] || 1013),
        windDir: weather.wind_direction_10m[idx] || 0,
        waveHeight: marine.wave_height?.[idx] || 0,
        rain: weather.precipitation[idx] || 0,
        cloud: weather.cloud_cover[idx] || 0
      };
      
      const scoreObj = calculateVidalisScore(dp, loc, moonIllum);
      const finalScore = Math.min(scoreObj.total / 100, 1);
      matrix[t].push(`${(finalScore*100).toFixed(0)}% ${getScoreEmoji(finalScore, dp.rain)}`);
      colors[t].push(getScoreColor(finalScore, dp.rain));
    }
  }
  return { matrix, colors };
}

function prepareForecastSheet(sheet) { 
  sheet.getRange("A4:Z200").clearContent().clearFormat(); 
}

function displayForecastResults(sheet, res) {
  sheet.getRange(5, 1, res.matrix.length, res.matrix[0].length).setValues(res.matrix);
  for(let r=0; r<res.colors.length; r++) {
    for(let c=1; c<res.colors[r].length; c++) {
      sheet.getRange(5+r, 1+c).setBackground(res.colors[r][c]);
    }
  }
}
