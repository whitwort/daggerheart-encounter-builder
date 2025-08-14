// Daggerheart Combat Encounter Automation System - v0.2

let CONFIG = {};
let BATTLE_VALUES = {};
let DAMAGE_TYPES = [];
let UI_CONFIG = {};
let GITHUB_CONFIG = {};

let CACHE_CONFIG = {};
let BATTLE_CALCULATION = {};
let GITHUB_RATE_LIMITS = {};
let FORM_VALIDATION = {};
let SHEET_FORMATTING = {};
let SEARCH_CONFIG = {};
let DIALOG_PRESETS = {};
let SYSTEM_TIMINGS = {};
let FEATURE_FLAGS = {};

/**
 * Enhanced multi-level caching system with automatic invalidation
 */
const DATA_CACHE = {
  // Core data caches
  adversaries: null,
  environments: null,
  filterOptions: null,
  
  // Configuration caches
  configuration: null,
  configLoadTime: 0,
  
  // File reference caches (avoid repeated folder searches)
  fileReferences: new Map(),
  folderReference: null,
  
  // Session timing
  lastAccessTime: 0,
  
  // Cache settings
  get CACHE_DURATION() {
    return (this.configuration?.CACHE_CONFIG?.CACHE_DURATION) || 300000;
  },
  
  get CONFIG_CACHE_DURATION() {
    return 600000; // 10 minutes for config
  },
  
  /**
   * Clear all cached data
   */
  clearAll() {
    this.adversaries = null;
    this.environments = null;
    this.filterOptions = null;
    this.configuration = null;
    this.configLoadTime = 0;
    this.fileReferences.clear();
    this.folderReference = null;
    this.lastAccessTime = 0;
  },
  
  /**
   * Clear only dynamic data (keep config and file refs)
   */
  clearData() {
    this.adversaries = null;
    this.environments = null;
    this.filterOptions = null;
    this.lastAccessTime = 0;
  },
  
  /**
   * Check if data cache is valid
   */
  isDataValid() {
    if (!this.adversaries || !this.environments) return false;
    const now = Date.now();
    return (now - this.lastAccessTime) < this.CACHE_DURATION;
  },
  
  /**
   * Check if configuration cache is valid
   */
  isConfigValid() {
    if (!this.configuration) return false;
    const now = Date.now();
    return (now - this.configLoadTime) < this.CONFIG_CACHE_DURATION;
  },
  
  /**
   * Get cached file reference or fetch and cache it
   */
  getFileReference(fileName) {
    if (this.fileReferences.has(fileName)) {
      return this.fileReferences.get(fileName);
    }
    
    // Get folder reference (cached)
    if (!this.folderReference) {
      this.folderReference = getOrCreateEncounterDataFolder();
    }
    
    // Find file and cache reference
    const files = this.folderReference.getFilesByName(fileName);
    const file = files.hasNext() ? files.next() : null;
    
    if (file) {
      this.fileReferences.set(fileName, file);
    }
    
    return file;
  }
};

/**
 * Lazy Configuration Manager
 * Loads configuration sections only when accessed
 */
const ConfigManager = {
  _rawConfig: null,
  _loaded: new Set(),
  _cache: {},
  
  /**
   * Initialize with minimal config loading
   */
  init() {
    // Only load critical configs needed for startup
    this.loadCriticalConfig();
  },
  
  /**
   * Load only the critical configuration needed at startup
   */
  loadCriticalConfig() {
    if (DATA_CACHE.isConfigValid() && DATA_CACHE.configuration) {
      this._rawConfig = DATA_CACHE.configuration;
      // Only assign critical values needed immediately
      CONFIG = this._rawConfig.CONFIG || {};
      BATTLE_VALUES = this._rawConfig.BATTLE_VALUES || {};
      DAMAGE_TYPES = this._rawConfig.DAMAGE_TYPES || [];
      return;
    }
    
    try {
      const configFile = DATA_CACHE.getFileReference('config.json');
      
      if (configFile) {
        const configContent = configFile.getBlob().getDataAsString();
        this._rawConfig = JSON.parse(configContent);
        
        // Cache the configuration
        DATA_CACHE.configuration = this._rawConfig;
        DATA_CACHE.configLoadTime = Date.now();
        
        // Load only critical sections
        CONFIG = this._rawConfig.CONFIG || {};
        BATTLE_VALUES = this._rawConfig.BATTLE_VALUES || {};
        DAMAGE_TYPES = this._rawConfig.DAMAGE_TYPES || [];
        
      } else {
        this.createDefaultConfiguration();
      }
    } catch (error) {
      console.error('Error loading critical configuration:', error);
      this.createDefaultConfiguration();
    }
  },
  
  /**
   * Lazy load a specific configuration section
   */
  getSection(sectionName) {
    // Return cached section if already loaded
    if (this._loaded.has(sectionName)) {
      return this._cache[sectionName];
    }
    
    // Ensure we have raw config
    if (!this._rawConfig) {
      this.loadCriticalConfig();
    }
    
    // Load and cache the section
    const sectionData = this._rawConfig[sectionName] || this.getDefaultSection(sectionName);
    this._cache[sectionName] = sectionData;
    this._loaded.add(sectionName);
    
    // Also assign to global variable for backward compatibility
    switch(sectionName) {
      case 'UI_CONFIG':
        UI_CONFIG = sectionData;
        break;
      case 'GITHUB_CONFIG':
        GITHUB_CONFIG = sectionData;
        break;
      case 'CACHE_CONFIG':
        CACHE_CONFIG = sectionData;
        break;
      case 'BATTLE_CALCULATION':
        BATTLE_CALCULATION = sectionData;
        break;
      case 'GITHUB_RATE_LIMITS':
        GITHUB_RATE_LIMITS = sectionData;
        break;
      case 'FORM_VALIDATION':
        FORM_VALIDATION = sectionData;
        break;
      case 'SHEET_FORMATTING':
        SHEET_FORMATTING = sectionData;
        break;
      case 'SEARCH_CONFIG':
        SEARCH_CONFIG = sectionData;
        break;
      case 'DIALOG_PRESETS':
        DIALOG_PRESETS = sectionData;
        break;
      case 'SYSTEM_TIMINGS':
        SYSTEM_TIMINGS = sectionData;
        break;
      case 'FEATURE_FLAGS':
        FEATURE_FLAGS = sectionData;
        break;
    }
    
    return sectionData;
  },
  
  /**
   * Get default values for a section
   */
  getDefaultSection(sectionName) {
    const defaults = {
      UI_CONFIG: {
        DIALOG_WIDTH: 1100,
        DIALOG_HEIGHT: 750,
        RESULTS_MAX_HEIGHT: 450,
        AUTO_CLOSE_DELAY: 10000,
        PROGRESS_SPINNER_DELAY: 1000
      },
      GITHUB_CONFIG: {
        REPO_BASE: 'https://api.github.com/repos/whitwort/daggerheart-srd/contents',
        ADVERSARIES_PATH: 'adversaries',
        ENVIRONMENTS_PATH: 'environments'
      },
      CACHE_CONFIG: {
        CACHE_DURATION: 300000,
        MIN_CACHE_CHECK_INTERVAL: 1000
      },
      BATTLE_CALCULATION: {
        BASE_MULTIPLIER: 3,
        BASE_ADDITION: 2,
        MULTIPLE_SOLOS_ADJUSTMENT: -2,
        LOWER_TIER_BONUS: 1,
        NO_ELITES_BONUS: 1,
        HIGH_DAMAGE_PENALTY: -2,
        MIN_SOLOS_FOR_ADJUSTMENT: 2,
        DEFAULT_PARTY_SIZE: 4,
        DEFAULT_PARTY_TIER: 2
      },
      GITHUB_RATE_LIMITS: {
        MAX_CONCURRENT_REQUESTS: 5,
        REQUEST_DELAY_MS: 200,
        RETRY_ATTEMPTS: 3
      },
      FORM_VALIDATION: {
        MIN_NAME_LENGTH: 1,
        MAX_NAME_LENGTH: 50,
        MIN_DIFFICULTY: 1,
        MAX_DIFFICULTY: 30,
        MIN_THRESHOLD_VALUE: 1,
        MAX_THRESHOLD_VALUE: 50,
        MIN_HP_VALUE: 1,
        MAX_HP_VALUE: 20,
        MIN_STRESS_VALUE: 1,
        MAX_STRESS_VALUE: 10,
        MIN_ATTACK_MODIFIER: -5,
        MAX_ATTACK_MODIFIER: 10,
        MIN_TIER: 1,
        MAX_TIER: 4
      },
      SHEET_FORMATTING: {
        NAME_COLUMN_WIDTH: 150,
        DIFFICULTY_COLUMN_WIDTH: 90,
        THRESHOLDS_COLUMN_WIDTH: 100,
        ATTACK_COLUMN_WIDTH: 80,
        STANDARD_ATTACK_COLUMN_WIDTH: 250,
        CHECKBOX_COLUMN_WIDTH: 45,
        NOTES_COLUMN_WIDTH: 200,
        MAX_COLUMNS_IN_SHEET: 12
      },
      SEARCH_CONFIG: {
        RESULTS_MAX_HEIGHT: 450,
        SEARCH_DEBOUNCE_DELAY: 300,
        MIN_SEARCH_CHARS: 1,
        MAX_SEARCH_RESULTS: 100
      },
      DIALOG_PRESETS: {
        ENCOUNTER_BUILDER: { WIDTH: 1100, HEIGHT: 750 },
        ADVERSARY_EDITOR: { WIDTH: 1000, HEIGHT: 800 },
        ENVIRONMENT_EDITOR: { WIDTH: 1000, HEIGHT: 800 },
        PROGRESS_DIALOG: { WIDTH: 500, HEIGHT: 300 },
        COMPLETION_DIALOG: { WIDTH: 600, HEIGHT: 400 }
      },
      SYSTEM_TIMINGS: {
        AUTO_CLOSE_DELAY: 10000,
        PROGRESS_SPINNER_DELAY: 1000,
        LOADING_ANIMATION_DURATION: 2000,
        ERROR_MESSAGE_TIMEOUT: 5000
      },
      FEATURE_FLAGS: {
        ENABLE_ADVANCED_SEARCH: true,
        ENABLE_BULK_OPERATIONS: false,
        ENABLE_EXPORT_FUNCTIONS: false,
        ENABLE_BACKUP_SYSTEM: false,
        ENABLE_DEBUG_LOGGING: false
      }
    };
    
    return defaults[sectionName] || {};
  },
  
  /**
   * Create default configuration (only when needed)
   */
  createDefaultConfiguration() {
    const defaultConfig = {
      CONFIG: this.getDefaultSection('CONFIG') || {
        ENCOUNTER_SHEET: 'Combat Encounters',
        TRACKER_SHEET: 'Health Tracker',
        ADVERSARY_DATA_SHEET: 'Adversary Templates',
        CUSTOM_ADVERSARIES_SHEET: 'Custom Adversaries'
      },
      BATTLE_VALUES: this.getDefaultSection('BATTLE_VALUES') || {
        'Minion': 0,
        'Standard': 2,
        'Horde': 2,
        'Skulk': 2,
        'Ranged': 2,
        'Support': 1,
        'Social': 1,
        'Leader': 3,
        'Bruiser': 4,
        'Solo': 5
      },
      DAMAGE_TYPES: [
        'physical', 'magic', 'psychic', 'poison', 'fire', 'ice', 'lightning'
      ]
    };
    
    this._rawConfig = defaultConfig;
    CONFIG = defaultConfig.CONFIG;
    BATTLE_VALUES = defaultConfig.BATTLE_VALUES;
    DAMAGE_TYPES = defaultConfig.DAMAGE_TYPES;
    
    // Save to file asynchronously
    try {
      const encounterDataFolder = getOrCreateEncounterDataFolder();
      const jsonContent = JSON.stringify({
        ...defaultConfig,
        ...Object.fromEntries(
          Object.keys(this.getDefaultSection).map(key => [key, this.getDefaultSection(key)])
        )
      }, null, 2);
      encounterDataFolder.createFile('config.json', jsonContent, MimeType.PLAIN_TEXT);
    } catch (error) {
      console.error('Error creating default configuration file:', error);
    }
  }
};


/**
 * Optimized loadConfiguration using lazy loading
 */
function loadConfiguration() {
  ConfigManager.init();
}

/**
 * Lazy wrapper functions for configuration access
 * These only load their specific sections when first called
 */
function getUIConfig() {
  if (!UI_CONFIG || Object.keys(UI_CONFIG).length === 0) {
    ConfigManager.getSection('UI_CONFIG');
  }
  return UI_CONFIG;
}

function getGitHubConfig() {
  if (!GITHUB_CONFIG || Object.keys(GITHUB_CONFIG).length === 0) {
    ConfigManager.getSection('GITHUB_CONFIG');
  }
  return GITHUB_CONFIG;
}

function getBattleCalculation() {
  if (!BATTLE_CALCULATION || Object.keys(BATTLE_CALCULATION).length === 0) {
    ConfigManager.getSection('BATTLE_CALCULATION');
  }
  return BATTLE_CALCULATION;
}

function getSheetFormatting() {
  if (!SHEET_FORMATTING || Object.keys(SHEET_FORMATTING).length === 0) {
    ConfigManager.getSection('SHEET_FORMATTING');
  }
  return SHEET_FORMATTING;
}

function getDialogPresets() {
  if (!DIALOG_PRESETS || Object.keys(DIALOG_PRESETS).length === 0) {
    ConfigManager.getSection('DIALOG_PRESETS');
  }
  return DIALOG_PRESETS;
}

function getFormValidation() {
  if (!FORM_VALIDATION || Object.keys(FORM_VALIDATION).length === 0) {
    ConfigManager.getSection('FORM_VALIDATION');
  }
  return FORM_VALIDATION;
}

function assignConfigurationValues(config) {
  CONFIG = config.CONFIG || {};
  BATTLE_VALUES = config.BATTLE_VALUES || {};
  DAMAGE_TYPES = config.DAMAGE_TYPES || [];
  UI_CONFIG = config.UI_CONFIG || {};
  GITHUB_CONFIG = config.GITHUB_CONFIG || {};
  CACHE_CONFIG = config.CACHE_CONFIG || {};
  BATTLE_CALCULATION = config.BATTLE_CALCULATION || {};
  GITHUB_RATE_LIMITS = config.GITHUB_RATE_LIMITS || {};
  FORM_VALIDATION = config.FORM_VALIDATION || {};
  SHEET_FORMATTING = config.SHEET_FORMATTING || {};
  SEARCH_CONFIG = config.SEARCH_CONFIG || {};
  DIALOG_PRESETS = config.DIALOG_PRESETS || {};
  SYSTEM_TIMINGS = config.SYSTEM_TIMINGS || {};
  FEATURE_FLAGS = config.FEATURE_FLAGS || {};
}

function createDefaultConfiguration() {
  const defaultConfig = {
    CONFIG: {
      ENCOUNTER_SHEET: 'Combat Encounters',
      TRACKER_SHEET: 'Health Tracker',
      ADVERSARY_DATA_SHEET: 'Adversary Templates',
      CUSTOM_ADVERSARIES_SHEET: 'Custom Adversaries'
    },
    BATTLE_VALUES: {
      'Minion': 0,
      'Standard': 2,
      'Horde': 2,
      'Skulk': 2,
      'Ranged': 2,
      'Support': 1,
      'Social': 1,
      'Leader': 3,
      'Bruiser': 4,
      'Solo': 5
    },
    DAMAGE_TYPES: [
      'physical', 'magic', 'psychic', 'poison', 'fire', 'ice', 'lightning'
    ],
    UI_CONFIG: {
      DIALOG_WIDTH: 1100,
      DIALOG_HEIGHT: 750,
      RESULTS_MAX_HEIGHT: 450,
      AUTO_CLOSE_DELAY: 10000,
      PROGRESS_SPINNER_DELAY: 1000
    },
    GITHUB_CONFIG: {
      REPO_BASE: 'https://api.github.com/repos/whitwort/daggerheart-srd/contents',
      ADVERSARIES_PATH: 'adversaries',
      ENVIRONMENTS_PATH: 'environments'
    },
    
    // Additional configuration sections
    CACHE_CONFIG: {
      CACHE_DURATION: 300000,
      MIN_CACHE_CHECK_INTERVAL: 1000
    },
    BATTLE_CALCULATION: {
      BASE_MULTIPLIER: 3,
      BASE_ADDITION: 2,
      MULTIPLE_SOLOS_ADJUSTMENT: -2,
      LOWER_TIER_BONUS: 1,
      NO_ELITES_BONUS: 1,
      HIGH_DAMAGE_PENALTY: -2,
      MIN_SOLOS_FOR_ADJUSTMENT: 2,
      DEFAULT_PARTY_SIZE: 4,
      DEFAULT_PARTY_TIER: 2
    },
    GITHUB_RATE_LIMITS: {
      MAX_CONCURRENT_REQUESTS: 5,
      REQUEST_DELAY_MS: 200,
      RETRY_ATTEMPTS: 3
    },
    FORM_VALIDATION: {
      MIN_NAME_LENGTH: 1,
      MAX_NAME_LENGTH: 50,
      MIN_DIFFICULTY: 1,
      MAX_DIFFICULTY: 30,
      MIN_THRESHOLD_VALUE: 1,
      MAX_THRESHOLD_VALUE: 50,
      MIN_HP_VALUE: 1,
      MAX_HP_VALUE: 20,
      MIN_STRESS_VALUE: 1,
      MAX_STRESS_VALUE: 10,
      MIN_ATTACK_MODIFIER: -5,
      MAX_ATTACK_MODIFIER: 10,
      MIN_TIER: 1,
      MAX_TIER: 4
    },
    SHEET_FORMATTING: {
      NAME_COLUMN_WIDTH: 150,
      DIFFICULTY_COLUMN_WIDTH: 90,
      THRESHOLDS_COLUMN_WIDTH: 100,
      ATTACK_COLUMN_WIDTH: 80,
      STANDARD_ATTACK_COLUMN_WIDTH: 250,
      CHECKBOX_COLUMN_WIDTH: 45,
      NOTES_COLUMN_WIDTH: 200,
      MAX_COLUMNS_IN_SHEET: 12
    },
    SEARCH_CONFIG: {
      RESULTS_MAX_HEIGHT: 450,
      SEARCH_DEBOUNCE_DELAY: 300,
      MIN_SEARCH_CHARS: 1,
      MAX_SEARCH_RESULTS: 100
    },
    DIALOG_PRESETS: {
      ENCOUNTER_BUILDER: { WIDTH: 1100, HEIGHT: 750 },
      ADVERSARY_EDITOR: { WIDTH: 1000, HEIGHT: 800 },
      ENVIRONMENT_EDITOR: { WIDTH: 1000, HEIGHT: 800 },
      PROGRESS_DIALOG: { WIDTH: 500, HEIGHT: 300 },
      COMPLETION_DIALOG: { WIDTH: 600, HEIGHT: 400 }
    },
    SYSTEM_TIMINGS: {
      AUTO_CLOSE_DELAY: 10000,
      PROGRESS_SPINNER_DELAY: 1000,
      LOADING_ANIMATION_DURATION: 2000,
      ERROR_MESSAGE_TIMEOUT: 5000
    },
    FEATURE_FLAGS: {
      ENABLE_ADVANCED_SEARCH: true,
      ENABLE_BULK_OPERATIONS: false,
      ENABLE_EXPORT_FUNCTIONS: false,
      ENABLE_BACKUP_SYSTEM: false,
      ENABLE_DEBUG_LOGGING: false
    }
  };
  
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(defaultConfig, null, 2);
    encounterDataFolder.createFile('config.json', jsonContent, MimeType.PLAIN_TEXT);
    
    console.log('Created default configuration file for v0.2');
  } catch (error) {
    console.error('Error creating default configuration file:', error);
  }
  
  // Always assign the configuration values, whether file creation succeeded or not
  assignConfigurationValues(defaultConfig);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daggerheart Combat')
    .addItem('Encounter Builder', 'showEncounterBuilder')
    .addSeparator()
    .addItem('Create/Edit Custom Adversary', 'showCustomAdversaryEditor')
    .addItem('Create/Edit Custom Environment', 'showCustomEnvironmentEditor')
    .addSeparator()
    .addItem('Import Core Adversaries from GitHub', 'fetchAndParseAdversariesFromGitHub')
    .addItem('Import Core Environments from GitHub', 'fetchAndParseEnvironmentsFromGitHub')
    .addSeparator()
    .addItem('Settings', 'showSettingsDialog')
    .addItem('Clear Data Cache', 'clearDataCache')
    .addToUi();
}

function clearDataCache() {
  DATA_CACHE.adversaries = null;
  DATA_CACHE.environments = null;
  DATA_CACHE.filterOptions = null;
  DATA_CACHE.lastAccessTime = 0;
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cache cleared successfully. Data will be reloaded on next use.');
}

function showEncounterBuilder() {
  loadConfiguration(); // Only loads critical config
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = createEncounterBuilderDialog();
  
  // Lazy load dialog presets only when needed
  const dialogSize = getDialogPresets().ENCOUNTER_BUILDER || { WIDTH: 1100, HEIGHT: 750 };
  htmlOutput.setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
  ui.showModalDialog(htmlOutput, 'Encounter Builder - v0.2');
}

function createEncounterBuilderDialog() {
  const html = generateOptimizedDialogHTML();
  return HtmlService.createHtmlOutput(html);
}

function getDialogData() {
  // Load configuration if needed (will use cache if valid)
  loadConfiguration();
  
  // Check if all data is cached and valid
  if (DATA_CACHE.isDataValid() && DATA_CACHE.adversaries && DATA_CACHE.environments && DATA_CACHE.filterOptions) {
    return {
      adversaries: DATA_CACHE.adversaries,
      environments: DATA_CACHE.environments,
      filterOptions: DATA_CACHE.filterOptions,
      battleValues: BATTLE_VALUES,
      damageTypes: DAMAGE_TYPES
    };
  }
  
  // Load data (will use cache if available)
  const adversaries = getAllAdversaryTemplates();
  const environments = getAllEnvironmentTemplates();
  
  // Only recalculate filter options if needed
  if (!DATA_CACHE.filterOptions || !DATA_CACHE.isDataValid()) {
    DATA_CACHE.filterOptions = extractFilterOptions(adversaries);
  }
  
  return {
    adversaries: adversaries,
    environments: environments,
    filterOptions: DATA_CACHE.filterOptions,
    battleValues: BATTLE_VALUES,
    damageTypes: DAMAGE_TYPES
  };
}

function extractFilterOptions(adversaries) {
  const values = Object.values(adversaries);
  
  const tiers = new Set();
  const types = new Set();
  const ranges = new Set();
  const resistances = new Set();
  
  values.forEach(adv => {
    // Add tier
    if (adv.tier) {
      tiers.add(adv.tier);
    }
    
    // Add type - filter out any falsy or empty values (FIXED BUG #2)
    if (adv.type) {
      const typeStr = String(adv.type).trim();
      if (typeStr.length > 0) {
        types.add(typeStr);
      }
    }
    
    // Add range
    if (adv.standardAttack && adv.standardAttack.range) {
      ranges.add(adv.standardAttack.range);
    }
    
    // IMPROVED: Enhanced resistance extraction from features
    if (adv.features) {
      adv.features.forEach(feature => {
        const desc = feature.description.toLowerCase();
        
        // Look for resistance patterns
        if (desc.includes('resistant to') || desc.includes('resistance to')) {
          // Extract what they're resistant to
          const resistanceMatch = desc.match(/resistant?\s+to\s+([^.]+)/i);
          if (resistanceMatch) {
            const resistanceText = resistanceMatch[1].toLowerCase();
            
            // Check for physical damage resistance
            if (resistanceText.includes('physical') || resistanceText.includes('phy')) {
              resistances.add('physical');
            }
            
            // Check for magic damage resistance  
            if (resistanceText.includes('magic') || resistanceText.includes('mag')) {
              resistances.add('magic');
            }
            
            // Also check for other potential damage types that might be added later
            if (resistanceText.includes('poison')) {
              resistances.add('poison');
            }
            if (resistanceText.includes('fire')) {
              resistances.add('fire');
            }
            if (resistanceText.includes('ice') || resistanceText.includes('cold')) {
              resistances.add('ice');
            }
            if (resistanceText.includes('lightning') || resistanceText.includes('electric')) {
              resistances.add('lightning');
            }
            if (resistanceText.includes('psychic') || resistanceText.includes('mental')) {
              resistances.add('psychic');
            }
          }
        }
        
        // Also check for immunity patterns
        if (desc.includes('immune to') || desc.includes('immunity to')) {
          const immunityMatch = desc.match(/immun[ei]\s+to\s+([^.]+)/i);
          if (immunityMatch) {
            const immunityText = immunityMatch[1].toLowerCase();
            
            if (immunityText.includes('physical') || immunityText.includes('phy')) {
              resistances.add('physical');
            }
            if (immunityText.includes('magic') || immunityText.includes('mag')) {
              resistances.add('magic');
            }
          }
        }
      });
    }
  });
  
  // Always include both core damage types since they exist in Daggerheart
  const damageTypes = ['physical', 'magic'];
  
  // Filter out any empty strings from types array (BUG FIX #2)
  const typesArray = Array.from(types).filter(t => t && t.length > 0);
  
  return {
    tiers: Array.from(tiers).sort((a, b) => a - b), // Sort numerically
    types: typesArray.sort(), // Filtered and sorted
    ranges: Array.from(ranges).sort(),
    resistances: Array.from(resistances).sort(),
    damageTypes: damageTypes // Always return both damage types
  };
}

function generateOptimizedDialogHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          ${getOptimizedDialogCSS()}
        </style>
      </head>
      <body>
        <div class="loading-overlay" id="loading-overlay">
          <div class="loading-spinner"></div>
          <div class="loading-text">Loading Encounter Builder...</div>
        </div>
        <div class="main-container" id="main-container" style="display: none;">
          <div class="left-panel" id="left-panel"></div>
          <div class="right-panel" id="right-panel"></div>
        </div>
        <script>
          ${generateOptimizedDialogScript()}
        </script>
      </body>
    </html>
  `;
}

function getOptimizedDialogCSS() {
  const maxHeight = SEARCH_CONFIG.RESULTS_MAX_HEIGHT || 450;
  
  return `
    * { box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 15px; 
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      min-height: 100vh;
    }
    .loading-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.95);
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    }
    .loading-spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #2196F3;
      border-radius: 50%;
      width: 50px;
      height: 50px;
      animation: spin 1s linear infinite;
      margin-bottom: 20px;
    }
    .loading-text {
      font-size: 18px;
      color: #333;
      font-weight: 600;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .main-container {
      display: flex;
      gap: 20px;
      max-width: 1400px;
      margin: 0 auto;
      height: calc(100vh - 30px);
    }
    .left-panel {
      flex: 2.2;
      background: white;
      padding: 25px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      height: 100%;
      overflow-y: auto;
    }
    .right-panel {
      flex: 1;
      background: white;
      padding: 25px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      height: 100%;
      overflow-y: auto;
      position: sticky;
      top: 15px;
    }
    .filter-section {
      background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
      padding: 25px;
      border-radius: 12px;
      margin-bottom: 25px;
      border: 2px solid #e3f2fd;
      box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    }
    .filter-header h3 {
      margin: 0 0 20px 0;
      color: #1976d2;
      font-size: 18px;
      font-weight: 600;
    }
    .filter-grid {
      display: flex;
      flex-direction: column;
      gap: 15px;
      margin-bottom: 20px;
    }
    .filter-row {
      display: flex;
      gap: 20px;
    }
    .filter-item {
      display: flex;
      flex-direction: column;
      gap: 8px;
      flex: 1;
    }
    .filter-label {
      font-weight: 600;
      color: #333;
      font-size: 14px;
    }
    select, input {
      padding: 12px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: all 0.3s ease;
      background: white;
    }
    select:focus, input:focus {
      border-color: #1976d2;
      outline: none;
      box-shadow: 0 0 0 3px rgba(25,118,210,0.2);
    }
    .difficulty-range {
      display: flex;
      gap: 12px;
      align-items: center;
    }
    .difficulty-range input { flex: 1; }
    .search-btn, .clear-btn {
      padding: 12px 28px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    .search-btn {
      background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
      color: white;
    }
    .clear-btn {
      background: linear-gradient(135deg, #757575 0%, #616161 100%);
      color: white;
    }
    .results-section {
      background: white;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      max-height: ${maxHeight}px;
      overflow-y: auto;
      overflow-x: hidden;
      word-wrap: break-word;
    }
    .results-count {
      text-align: center;
      padding: 20px;
      font-weight: 600;
      color: #555;
      background: #fafafa;
      position: sticky;
      top: 0;
      z-index: 10;
    }
    .result-item {
      padding: 15px;
      border-bottom: 1px solid #f0f0f0;
      cursor: pointer;
      transition: all 0.3s ease;
      position: relative;
    }
    .result-item:hover {
      background: #f8f9fa;
      transform: translateX(4px);
    }
    .result-item.selected {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border-left: 5px solid #4caf50;
      transform: translateX(8px);
    }
    .result-name {
      font-weight: 700;
      font-size: 16px;
      color: #1976d2;
      margin-bottom: 5px;
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 8px;
    }
    .result-badges {
      display: flex;
      align-items: center;
      gap: 4px;
      flex-shrink: 0;
    }
    .result-stats {
      font-size: 13px;
      color: #666;
      margin-bottom: 8px;
      line-height: 1.4;
    }
    .result-description {
      font-size: 12px;
      color: #444;
      line-height: 1.5;
      font-style: italic;
    }
    .selection-badge {
      background: #4caf50;
      color: white;
      padding: 4px 8px;
      border-radius: 12px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .tier-badge {
      background: #ff9800;
      color: white;
      padding: 2px 6px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .type-badge {
      background: #9c27b0;
      color: white;
      padding: 2px 6px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .no-results {
      text-align: center;
      padding: 60px 40px;
      color: #999;
      font-style: italic;
      font-size: 16px;
    }
    .encounter-builder {
      height: 100%;
      display: flex;
      flex-direction: column;
    }
    .builder-header {
      text-align: center;
      margin-bottom: 25px;
      padding-bottom: 20px;
      border-bottom: 3px solid #e3f2fd;
    }
    .builder-header h3 {
      margin: 0;
      color: #1976d2;
      font-size: 22px;
      font-weight: 700;
    }
    .config-section {
      background: #f5f5f5;
      border: 2px solid #bdbdbd;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .config-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 15px;
    }
    .config-row:last-child { margin-bottom: 0; }
    .player-count-input, .player-tier-select {
      width: 80px;
      padding: 8px;
      font-size: 14px;
      text-align: center;
      border: 2px solid #1976d2;
      border-radius: 6px;
    }
    .high-damage-checkbox {
      transform: scale(1.3);
      accent-color: #1976d2;
    }
    .battle-points {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border: 3px solid #4caf50;
      border-radius: 16px;
      padding: 20px;
      margin-bottom: 20px;
      text-align: center;
    }
    .points-display {
      font-size: 24px;
      font-weight: 700;
      color: #2e7d32;
      margin-bottom: 12px;
    }
    .points-breakdown {
      font-size: 13px;
      color: #555;
      line-height: 1.4;
      margin-bottom: 10px;
    }
    .difficulty-breakdown {
      font-size: 11px;
      color: #666;
      margin-top: 6px;
      font-style: italic;
      line-height: 1.3;
    }
    .difficulty-indicator {
      margin-top: 12px;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: 700;
      display: inline-block;
      font-size: 14px;
    }
    .easy { background: #c8e6c9; color: #1b5e20; }
    .normal { background: #fff3c4; color: #000000; }
    .hard { background: #ffcdd2; color: #b71c1c; }
    .deadly { background: #f3e5f5; color: #4a148c; }
    .selected-section {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border: 3px solid #4caf50;
      border-radius: 16px;
      margin-bottom: 20px;
      padding: 20px;
      flex: 1;
      display: flex;
      flex-direction: column;
    }
    .selected-header {
      font-weight: 700;
      color: #2e7d32;
      margin-bottom: 15px;
      font-size: 18px;
      text-align: center;
    }
    .selected-list {
      flex: 1;
      overflow-y: auto;
      margin-bottom: 15px;
    }
    .selected-adversary {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px;
      background: white;
      border: 2px solid #a5d6a7;
      border-radius: 10px;
      margin-bottom: 10px;
      font-size: 13px;
    }
    .selected-name {
      font-weight: 700;
      color: #2e7d32;
      flex: 1;
    }
    .selected-stats {
      font-size: 11px;
      color: #666;
      margin-top: 2px;
    }
    .selected-count {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .count-btn {
      background: #4caf50;
      color: white;
      border: none;
      width: 24px;
      height: 24px;
      border-radius: 50%;
      cursor: pointer;
      font-weight: bold;
      font-size: 14px;
    }
    .count-display {
      min-width: 20px;
      text-align: center;
      font-weight: bold;
      font-size: 16px;
      color: #2e7d32;
    }
    .remove-btn {
      background: #f44336;
      color: white;
      border: none;
      padding: 4px 8px;
      border-radius: 6px;
      cursor: pointer;
      font-size: 12px;
      font-weight: bold;
    }
    .build-encounter-btn {
      background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
      color: white;
      border: none;
      padding: 16px 24px;
      border-radius: 12px;
      font-size: 16px;
      font-weight: 700;
      cursor: pointer;
      width: 100%;
      transition: all 0.3s ease;
    }
    .build-encounter-btn:disabled {
      background: #bdbdbd;
      cursor: not-allowed;
    }
    .version-info {
      text-align: center;
      font-size: 11px;
      color: #999;
      margin-top: 15px;
      font-style: italic;
    }
    @media (max-width: 1200px) {
      .main-container { flex-direction: column; height: auto; }
      .right-panel { position: static; margin-top: 20px; }
      .filter-grid { grid-template-columns: 1fr; }
    }
  `;
}

function generateOptimizedDialogScript() {
  // Lazy load configuration values and pass them to client script
  const config = getBattleCalculation();
  const defaultPartySize = config.DEFAULT_PARTY_SIZE || 4;
  const defaultPartyTier = config.DEFAULT_PARTY_TIER || 2;
  
  return `
    // ============================================
    // Enhanced DOM Management System
    // ============================================
    
    let adversariesData = {};
    let environmentsData = {};
    let battleValues = {};
    let damageTypes = [];
    let filterOptions = {};
    let selectedAdversaries = {};
    let lastSearchResults = [];
    let dataLoaded = false;
    
    // Battle calculation configuration - loaded from server
    const BATTLE_CONFIG = {
      BASE_MULTIPLIER: ${config.BASE_MULTIPLIER || 3},
      BASE_ADDITION: ${config.BASE_ADDITION || 2},
      MULTIPLE_SOLOS_ADJUSTMENT: ${config.MULTIPLE_SOLOS_ADJUSTMENT || -2},
      LOWER_TIER_BONUS: ${config.LOWER_TIER_BONUS || 1},
      NO_ELITES_BONUS: ${config.NO_ELITES_BONUS || 1},
      HIGH_DAMAGE_PENALTY: ${config.HIGH_DAMAGE_PENALTY || -2},
      MIN_SOLOS_FOR_ADJUSTMENT: ${config.MIN_SOLOS_FOR_ADJUSTMENT || 2},
      DEFAULT_PARTY_SIZE: ${defaultPartySize},
      DEFAULT_PARTY_TIER: ${defaultPartyTier}
    };
    
    // DOM element cache to avoid repeated queries
    const DOM_CACHE = {
      elements: new Map(),
      
      get(id) {
        if (!this.elements.has(id)) {
          this.elements.set(id, document.getElementById(id));
        }
        return this.elements.get(id);
      },
      
      clear() {
        this.elements.clear();
      }
    };
    
    // Virtual DOM for search results
    const VDOM = {
      currentResults: new Map(),
      resultContainer: null,
      resultPool: [], // Reusable DOM elements
      maxPoolSize: 100,
      
      init() {
        this.resultContainer = DOM_CACHE.get('results-section');
      },
      
      // Get or create a result item element
      getResultElement() {
        if (this.resultPool.length > 0) {
          return this.resultPool.pop();
        }
        
        const div = document.createElement('div');
        div.className = 'result-item';
        return div;
      },
      
      // Return element to pool for reuse
      releaseElement(element) {
        if (this.resultPool.length < this.maxPoolSize) {
          element.className = 'result-item';
          element.onclick = null;
          this.resultPool.push(element);
        }
      },
      
      // Efficiently update only changed results
      updateResults(results) {
        const newResultsMap = new Map();
        const fragment = document.createDocumentFragment();
        let hasChanges = false;
        
        // Check if we need full rebuild (different result count or first render)
        if (this.currentResults.size === 0 || this.currentResults.size !== results.length) {
          hasChanges = true;
          this.renderFullResults(results);
          return;
        }
        
        // Incremental update for same-size result sets
        const container = this.resultContainer;
        const existingNodes = container.querySelectorAll('.result-item');
        
        results.forEach((adv, index) => {
          const existingNode = existingNodes[index];
          if (existingNode) {
            const oldName = existingNode.dataset.advName;
            if (oldName !== adv.name) {
              // Different adversary, update the node
              this.updateResultNode(existingNode, adv);
              hasChanges = true;
            } else {
              // Same adversary, just update selection state
              const isSelected = selectedAdversaries[adv.name];
              const wasSelected = existingNode.classList.contains('selected');
              if (isSelected !== wasSelected) {
                existingNode.classList.toggle('selected', isSelected);
                this.updateResultBadges(existingNode, adv);
                hasChanges = true;
              }
            }
          }
          newResultsMap.set(adv.name, adv);
        });
        
        this.currentResults = newResultsMap;
      },
      
      // Full render when structure changes significantly
      renderFullResults(results) {
        const container = this.resultContainer;
        
        // Return all current elements to pool
        const existingItems = container.querySelectorAll('.result-item');
        existingItems.forEach(item => this.releaseElement(item));
        
        // Use DocumentFragment for batch DOM operations
        const fragment = document.createDocumentFragment();
        
        // Add count header
        const countDiv = document.createElement('div');
        countDiv.className = 'results-count';
        countDiv.textContent = results.length + ' adversar' + (results.length === 1 ? 'y' : 'ies') + ' found';
        fragment.appendChild(countDiv);
        
        // Add result items
        results.forEach(adv => {
          const div = this.getResultElement();
          this.updateResultNode(div, adv);
          fragment.appendChild(div);
        });
        
        // Single DOM update
        container.innerHTML = '';
        container.appendChild(fragment);
        
        // Update cache
        this.currentResults.clear();
        results.forEach(adv => this.currentResults.set(adv.name, adv));
      },
      
      // Update a single result node
      updateResultNode(node, adv) {
        node.dataset.advName = adv.name;
        // FIX: Check if the adversary is actually in selectedAdversaries object
        const isSelected = adv.name in selectedAdversaries && selectedAdversaries[adv.name] > 0;
        node.className = 'result-item' + (isSelected ? ' selected' : '');
        node.onclick = () => toggleAdversarySelection(adv.name);
        
        // Use innerHTML once for better performance
        node.innerHTML = this.getResultHTML(adv);
      },
      
      // Update only badges without rebuilding entire node
      updateResultBadges(node, adv) {
        const badgesContainer = node.querySelector('.result-badges');
        if (badgesContainer) {
          badgesContainer.innerHTML = this.getBadgesHTML(adv);
        }
      },
      
      // Generate HTML strings (faster than DOM manipulation)
      getResultHTML(adv) {
        return \`
          <div class="result-name">
            <span>\${adv.name}</span>
            <div class="result-badges">\${this.getBadgesHTML(adv)}</div>
          </div>
          <div class="result-stats">\${this.getStatsHTML(adv)}</div>
          <div class="result-description">\${adv.description}</div>
        \`;
      },
      
      getBadgesHTML(adv) {
        let html = \`<span class="tier-badge">T\${adv.tier}</span>\`;
        html += \`<span class="type-badge">\${adv.type}</span>\`;
        if (selectedAdversaries[adv.name]) {
          html += \`<span class="selection-badge">×\${selectedAdversaries[adv.name]}</span>\`;
        }
        return html;
      },
      
      getStatsHTML(adv) {
        return \`<strong>Difficulty:</strong> \${adv.difficulty} | <strong>HP:</strong> \${adv.hp} | <strong>Stress:</strong> \${adv.stress} | <strong>Value:</strong> \${getBattleValueJS(adv.type)}<br>
                <strong>Attack:</strong> +\${adv.attackModifier} \${adv.standardAttack.name} (\${adv.standardAttack.range}, \${adv.standardAttack.damage})\`;
      }
    };
    
    // Debounced search to reduce unnecessary updates
    const SearchManager = {
      searchTimeout: null,
      lastQuery: '',
      
      debounceSearch() {
        if (this.searchTimeout) {
          clearTimeout(this.searchTimeout);
        }
        
        this.searchTimeout = setTimeout(() => {
          this.executeSearch();
        }, 150); // Reduced from 300ms for snappier feel
      },
      
      executeSearch() {
        if (!dataLoaded) return;
        
        const filters = this.getFilterValues();
        const queryKey = JSON.stringify(filters);
        
        // Skip if query hasn't changed
        if (queryKey === this.lastQuery) return;
        this.lastQuery = queryKey;
        
        let results = Object.values(adversariesData);
        results = this.applyFilters(results, filters);
        results.sort((a, b) => a.name.localeCompare(b.name));
        
        lastSearchResults = results;
        displayResults(results);
      },
      
      getFilterValues() {
        return {
          tier: DOM_CACHE.get('tier-filter').value,
          type: DOM_CACHE.get('type-filter').value,
          range: DOM_CACHE.get('range-filter').value,
          difficultyMin: DOM_CACHE.get('difficulty-min').value,
          difficultyMax: DOM_CACHE.get('difficulty-max').value,
          damageType: DOM_CACHE.get('damage-type-filter').value,
          resistance: DOM_CACHE.get('resistance-filter').value,
          textSearch: DOM_CACHE.get('text-search').value.toLowerCase(),
          includeCore: DOM_CACHE.get('include-core').checked,
          includeCustom: DOM_CACHE.get('include-custom').checked
        };
      },
      
      applyFilters(results, filters) {
        // Source filter first (most restrictive)
        if (!filters.includeCore || !filters.includeCustom) {
          results = results.filter(adv => {
            const isCustom = adv.isCustom === true;
            return (filters.includeCore || isCustom) && (filters.includeCustom || !isCustom);
          });
        }
        
        // Apply other filters only if set
        if (filters.tier) {
          results = results.filter(adv => adv.tier == filters.tier);
        }
        if (filters.type) {
          results = results.filter(adv => adv.type === filters.type);
        }
        if (filters.range) {
          results = results.filter(adv => adv.standardAttack.range === filters.range);
        }
        if (filters.difficultyMin) {
          const min = parseInt(filters.difficultyMin);
          results = results.filter(adv => adv.difficulty >= min);
        }
        if (filters.difficultyMax) {
          const max = parseInt(filters.difficultyMax);
          results = results.filter(adv => adv.difficulty <= max);
        }
        
        // Text-based filters (more expensive)
        if (filters.damageType) {
          results = results.filter(adv => {
            if (!adv.standardAttack || !adv.standardAttack.damage) return false;
            const damageStr = adv.standardAttack.damage.toLowerCase().trim();
            
            if (filters.damageType === 'physical') {
              // If it contains "physical" or "phy", it's physical
              // If it has NO damage type specified (just dice), assume physical per Daggerheart rules
              const hasPhysical = damageStr.includes('physical') || damageStr.includes('phy');
              const hasMagic = damageStr.includes('magic') || damageStr.includes('mag');
              
              return hasPhysical || (!hasPhysical && !hasMagic);
            } else if (filters.damageType === 'magic') {
              // Only include if explicitly says magic/mag
              return damageStr.includes('magic') || damageStr.includes('mag');
            }
            return false;
          });
        }
        
        if (filters.resistance) {
          results = results.filter(adv => {
            if (!adv.features) return false;
            return adv.features.some(f => {
              const desc = f.description.toLowerCase();
              return desc.includes('resistant to ' + filters.resistance) || 
                     desc.includes('resistance to ' + filters.resistance);
            });
          });
        }
        
        // Full text search (most expensive, do last)
        if (filters.textSearch) {
          results = results.filter(adv => {
            // Cache searchable text on adversary object
            if (!adv._searchCache) {
              adv._searchCache = [
                adv.name, adv.description, adv.motives,
                adv.standardAttack.name, adv.standardAttack.damage,
                ...adv.experience,
                ...adv.features.map(f => f.name + ' ' + f.description)
              ].join(' ').toLowerCase();
            }
            return adv._searchCache.includes(filters.textSearch);
          });
        }
        
        return results;
      }
    };
    
    // Optimized initialization
    function initializeDialog() {
      google.script.run
        .withSuccessHandler(onDataLoaded)
        .withFailureHandler(onDataLoadError)
        .getDialogData();
    }

    function onDataLoaded(data) {
      adversariesData = data.adversaries;
      environmentsData = data.environments;
      battleValues = data.battleValues;
      damageTypes = data.damageTypes;
      filterOptions = data.filterOptions;
      dataLoaded = true;
      
      buildUI();
      VDOM.init();
      
      DOM_CACHE.get('loading-overlay').style.display = 'none';
      DOM_CACHE.get('main-container').style.display = 'flex';
      
      // Use requestAnimationFrame for smooth initial render
      requestAnimationFrame(() => {
        initializeAllAdversaries();
        updateBattlePoints();
      });
    }

    function onDataLoadError(error) {
      console.error('Error loading data:', error);
      DOM_CACHE.get('loading-overlay').innerHTML = 
        '<div style="color: red;">Error loading data. Please try again.</div>';
    }

    function buildUI() {
      const leftPanel = DOM_CACHE.get('left-panel');
      const rightPanel = DOM_CACHE.get('right-panel');
      
      leftPanel.innerHTML = generateLeftPanelHTML();
      rightPanel.innerHTML = generateRightPanelHTML();
      
      // Cache frequently used elements
      DOM_CACHE.clear();
    }

    function generateLeftPanelHTML() {
      const html = [];
      html.push('<div class="filter-section">');
      html.push('<div class="filter-header"><h3>Adversary Filters</h3></div>');
      html.push('<div class="filter-grid">');
      
      html.push('<div class="filter-row">');
      html.push(createFilterItem('Tier', 'tier-filter', 'Any Tier', 
        filterOptions.tiers.map(t => ({value: t, label: 'Tier ' + t}))));
      html.push(createFilterItem('Type', 'type-filter', 'Any Type',
        filterOptions.types.map(t => ({value: t, label: t}))));
      html.push(createDifficultyRangeFilter());
      html.push('</div>');
      
      html.push('<div class="filter-row">');
      html.push(createFilterItem('Range', 'range-filter', 'Any Range',
        filterOptions.ranges.map(r => ({value: r, label: r}))));
      html.push(createFilterItem('Damage Type', 'damage-type-filter', 'Any Damage Type',
        damageTypes.map(t => ({value: t, label: t.charAt(0).toUpperCase() + t.slice(1)}))));
      html.push(createFilterItem('Damage Resistance', 'resistance-filter', 'Any Resistances',
        filterOptions.resistances.map(r => ({value: r, label: r.charAt(0).toUpperCase() + r.slice(1)}))));
      html.push('</div>');
      
      html.push('<div class="filter-row">');
      html.push(createTextSearchFilter());
      html.push(createSourceFilter());
      html.push(createClearButton());
      html.push('</div>');
      
      html.push('</div></div>');
      html.push('<div class="results-section" id="results-section">');
      html.push('<div class="no-results">Loading adversaries...</div>');
      html.push('</div>');
      
      return html.join('');
    }

    function createFilterItem(label, id, defaultText, options) {
      const html = [];
      html.push('<div class="filter-item">');
      html.push('<label class="filter-label">' + label + '</label>');
      html.push('<select id="' + id + '" onchange="searchAdversaries()">');
      html.push('<option value="">' + defaultText + '</option>');
      options.forEach(opt => {
        html.push('<option value="' + opt.value + '">' + opt.label + '</option>');
      });
      html.push('</select></div>');
      return html.join('');
    }

    function createDifficultyRangeFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Difficulty Range</label>' +
        '<div class="difficulty-range">' +
        '<input type="number" id="difficulty-min" placeholder="Min" min="1" max="30" onchange="searchAdversaries()">' +
        '<span>to</span>' +
        '<input type="number" id="difficulty-max" placeholder="Max" min="1" max="30" onchange="searchAdversaries()">' +
        '</div></div>';
    }

    function createTextSearchFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Text Search</label>' +
        '<input type="text" id="text-search" placeholder="Search names, descriptions, features..." oninput="searchAdversaries()">' +
        '</div>';
    }

    function createSourceFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Adversary Sources</label>' +
        '<div style="display: flex; gap: 15px; align-items: center;">' +
        '<label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">' +
        '<input type="checkbox" id="include-core" checked onchange="searchAdversaries()" style="transform: scale(1.2);">' +
        'Core</label>' +
        '<label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">' +
        '<input type="checkbox" id="include-custom" checked onchange="searchAdversaries()" style="transform: scale(1.2);">' +
        'Custom</label></div></div>';
    }

    function createClearButton() {
      return '<div class="filter-item" style="display: flex; flex-direction: column;">' +
        '<label class="filter-label" style="visibility: hidden;">Clear</label>' +
        '<button class="clear-btn" onclick="clearFilters()">Clear Filters</button>' +
        '</div>';
    }

    function generateRightPanelHTML() {
      const html = [];
      html.push('<div class="encounter-builder">');
      html.push('<div class="builder-header"><h3>Encounter Builder</h3></div>');
      
      html.push('<div class="config-section">');
      html.push(createConfigRow('Players:', 'player-count', 'number', '${defaultPartySize}', 'player-count-input'));
      html.push(createTierSelect());
      html.push(createHighDamageCheckbox());
      html.push(createEnvironmentSelect());
      html.push('</div>');
      
      html.push('<div class="battle-points">');
      html.push('<div class="points-display" id="total-points">0 Battle Points</div>');
      html.push('<div class="points-breakdown" id="points-breakdown">Select adversaries to see difficulty</div>');
      html.push('<div class="difficulty-breakdown" id="difficulty-breakdown">Configure encounter to see calculation details</div>');
      html.push('<div class="difficulty-indicator" id="difficulty-indicator" style="display: none;"></div>');
      html.push('</div>');
      
      html.push('<div class="selected-section" id="selected-section" style="display: none;">');
      html.push('<div class="selected-header">Selected Adversaries</div>');
      html.push('<div class="selected-list" id="selected-list"></div>');
      html.push('<button class="build-encounter-btn" id="build-btn" onclick="buildEncounterFromSelected()">Build Encounter</button>');
      html.push('</div>');
      
      html.push('<div class="version-info">Daggerheart Combat System v0.1</div>');
      html.push('</div>');
      
      return html.join('');
    }

    function createConfigRow(label, id, type, value, className) {
      return '<div class="config-row">' +
        '<label><strong>' + label + '</strong></label>' +
        '<input type="' + type + '" id="' + id + '" class="' + className + '" value="' + value + '" min="1" max="8" onchange="updateBattlePoints()">' +
        '</div>';
    }

    function createTierSelect() {
      const defaultTier = ${defaultPartyTier};
      return '<div class="config-row">' +
        '<label><strong>Tier:</strong></label>' +
        '<select id="player-tier" class="player-tier-select" onchange="updateBattlePoints()">' +
        '<option value="1"' + (defaultTier === 1 ? ' selected' : '') + '>Tier 1</option>' +
        '<option value="2"' + (defaultTier === 2 ? ' selected' : '') + '>Tier 2</option>' +
        '<option value="3"' + (defaultTier === 3 ? ' selected' : '') + '>Tier 3</option>' +
        '<option value="4"' + (defaultTier === 4 ? ' selected' : '') + '>Tier 4</option>' +
        '</select></div>';
    }

    function createHighDamageCheckbox() {
      return '<div class="config-row">' +
        '<label><strong>High Damage:</strong></label>' +
        '<input type="checkbox" id="high-damage" class="high-damage-checkbox" onchange="updateBattlePoints()">' +
        '</div>';
    }

    function createEnvironmentSelect() {
      const envNames = Object.keys(environmentsData).sort();
      const options = envNames.map(name => '<option value="' + name + '">' + name + '</option>').join('');
      return '<div class="config-row">' +
        '<label><strong>Environment:</strong></label>' +
        '<select id="environment-select" class="player-tier-select" style="width: 200px;">' +
        '<option value="">No Environment</option>' + options +
        '</select></div>';
    }

    function getBattleValueJS(adversaryType) {
      return battleValues[adversaryType] || 2;
    }

    function initializeAllAdversaries() {
      if (!dataLoaded) return;
      let results = Object.values(adversariesData);
      results.sort((a, b) => a.name.localeCompare(b.name));
      lastSearchResults = results;
      displayResults(results);
    }

    // Use new optimized search
    function searchAdversaries() {
      SearchManager.debounceSearch();
    }

    // Use new optimized display
    function displayResults(results) {
      if (results.length === 0) {
        VDOM.resultContainer.innerHTML = '<div class="no-results">No adversaries found matching your criteria</div>';
        VDOM.currentResults.clear();
        return;
      }
      
      VDOM.updateResults(results);
    }
    
    function toggleAdversarySelection(name) {
      if (selectedAdversaries[name]) {
        delete selectedAdversaries[name];
      } else {
        selectedAdversaries[name] = 1;
      }
      updateSelectedDisplay();
      updateBattlePoints();
      
      // Optimized: only update the specific result item
      const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
      resultItems.forEach(item => {
        if (item.dataset.advName === name) {
          item.classList.toggle('selected', selectedAdversaries[name]);
          VDOM.updateResultBadges(item, adversariesData[name]);
        }
      });
    }
    
    // Optimized selected display with fragment
    function updateSelectedDisplay() {
      const section = DOM_CACHE.get('selected-section');
      const list = DOM_CACHE.get('selected-list');
      
      const names = Object.keys(selectedAdversaries);
      
      if (names.length === 0) {
        section.style.display = 'none';
        return;
      }
      
      section.style.display = 'block';
      
      // Use DocumentFragment for batch updates
      const fragment = document.createDocumentFragment();
      
      names.forEach(name => {
        const adv = adversariesData[name];
        const count = selectedAdversaries[name];
        
        const div = document.createElement('div');
        div.className = 'selected-adversary';
        div.innerHTML = createSelectedAdversaryHTML(name, adv, count);
        fragment.appendChild(div);
      });
      
      list.innerHTML = '';
      list.appendChild(fragment);
    }

    function createSelectedAdversaryHTML(name, adv, count) {
      return '<div style="flex: 1;">' +
        '<div class="selected-name">' + name + '</div>' +
        '<div class="selected-stats">T' + adv.tier + ' ' + adv.type + 
        ' | Diff: ' + adv.difficulty + ' | Value: ' + getBattleValueJS(adv.type) + '</div>' +
        '</div>' +
        '<div class="selected-count">' +
        '<button class="count-btn" onclick="changeCount(\\'' + name.replace(/'/g, "\\'") + '\\', -1)">-</button>' +
        '<span class="count-display">' + count + '</span>' +
        '<button class="count-btn" onclick="changeCount(\\'' + name.replace(/'/g, "\\'") + '\\', 1)">+</button>' +
        '<button class="remove-btn" onclick="removeSelected(\\'' + name.replace(/'/g, "\\'") + '\\')">×</button>' +
        '</div>';
    }

    // Cached battle points calculation
    let lastBattlePointsState = '';
    let lastBattlePointsResult = null;
    
    function updateBattlePoints() {
      if (!dataLoaded) return;
      
      try {
        // Create state key to check if recalculation needed
        const stateKey = JSON.stringify({
          selected: selectedAdversaries,
          playerCount: DOM_CACHE.get('player-count').value,
          playerTier: DOM_CACHE.get('player-tier').value,
          highDamage: DOM_CACHE.get('high-damage').checked
        });
        
        // Use cached result if nothing changed
        if (stateKey === lastBattlePointsState && lastBattlePointsResult) {
          applyBattlePointsDisplay(lastBattlePointsResult);
          return;
        }
        
        let totalPoints = 0;
        const breakdown = [];
        const selectedList = [];
        
        const playerCount = parseInt(DOM_CACHE.get('player-count').value) || BATTLE_CONFIG.DEFAULT_PARTY_SIZE;
        
        Object.keys(selectedAdversaries).forEach(name => {
          const count = selectedAdversaries[name];
          if (adversariesData[name] && count > 0) {
            const template = adversariesData[name];
            let points;
            
            if (template.type === 'Minion') {
              const groups = Math.ceil(count / playerCount);
              points = groups;
              breakdown.push(count + 'x ' + name + ' (' + groups + ' group' + 
                            (groups > 1 ? 's' : '') + ' = ' + points + ' pts)');
            } else {
              const value = getBattleValueJS(template.type);
              points = value * count;
              breakdown.push(count + 'x ' + name + ' (' + points + ' pts)');
            }
            
            totalPoints += points;
            
            for (let i = 0; i < count; i++) {
              selectedList.push(template);
            }
          }
        });
        
        const result = {
          totalPoints: totalPoints,
          breakdown: breakdown,
          selectedList: selectedList,
          playerCount: playerCount,
          playerTier: parseInt(DOM_CACHE.get('player-tier').value) || BATTLE_CONFIG.DEFAULT_PARTY_TIER,
          isHighDamage: DOM_CACHE.get('high-damage').checked
        };
        
        // Cache the result
        lastBattlePointsState = stateKey;
        lastBattlePointsResult = result;
        
        applyBattlePointsDisplay(result);
        
      } catch (error) {
        console.error('Error in updateBattlePoints:', error);
      }
    }
    
    function applyBattlePointsDisplay(result) {
      DOM_CACHE.get('total-points').textContent = result.totalPoints + ' Battle Points';
      
      if (result.breakdown.length > 0) {
        DOM_CACHE.get('points-breakdown').innerHTML = result.breakdown.join('<br>');
        
        const difficulty = getDifficultyLevel(
          result.totalPoints, 
          result.playerCount, 
          result.playerTier, 
          result.selectedList, 
          result.isHighDamage
        );
        
        const indicator = DOM_CACHE.get('difficulty-indicator');
        indicator.textContent = difficulty.label;
        indicator.className = 'difficulty-indicator ' + difficulty.class;
        indicator.style.display = 'inline-block';
        
        DOM_CACHE.get('difficulty-breakdown').innerHTML = difficulty.breakdown;
        DOM_CACHE.get('build-btn').disabled = false;
      } else {
        DOM_CACHE.get('points-breakdown').textContent = 'Select adversaries to see difficulty';
        DOM_CACHE.get('difficulty-breakdown').textContent = 'Configure encounter to see calculation details';
        DOM_CACHE.get('difficulty-indicator').style.display = 'none';
        DOM_CACHE.get('build-btn').disabled = true;
      }
    }

    function getDifficultyLevel(points, playerCount, playerTier, adversaries, isHighDamage) {
      const originalBase = (BATTLE_CONFIG.BASE_MULTIPLIER * playerCount) + BATTLE_CONFIG.BASE_ADDITION;
      let adjustedBase = originalBase;
      let adjustments = [];
      
      const soloCount = adversaries.filter(adv => adv.type === 'Solo').length;
      if (soloCount >= BATTLE_CONFIG.MIN_SOLOS_FOR_ADJUSTMENT) {
        adjustedBase += BATTLE_CONFIG.MULTIPLE_SOLOS_ADJUSTMENT;
        adjustments.push(BATTLE_CONFIG.MULTIPLE_SOLOS_ADJUSTMENT + ' (multiple solos)');
      }
      
      const hasLowerTier = adversaries.some(adv => adv.tier < playerTier);
      if (hasLowerTier) {
        adjustedBase += BATTLE_CONFIG.LOWER_TIER_BONUS;
        adjustments.push('+' + BATTLE_CONFIG.LOWER_TIER_BONUS + ' (lower tier adversaries)');
      }
      
      const hasElites = adversaries.some(adv => 
        ['Bruiser', 'Horde', 'Leader', 'Solo'].includes(adv.type)
      );
      if (!hasElites && adversaries.length > 0) {
        adjustedBase += BATTLE_CONFIG.NO_ELITES_BONUS;
        adjustments.push('+' + BATTLE_CONFIG.NO_ELITES_BONUS + ' (no elite types)');
      }
      
      if (isHighDamage) {
        adjustedBase += BATTLE_CONFIG.HIGH_DAMAGE_PENALTY;
        adjustments.push(BATTLE_CONFIG.HIGH_DAMAGE_PENALTY + ' (high damage encounter)');
      }
      
      const easyThreshold = adjustedBase - 1;
      const normalMax = adjustedBase;
      const hardMax = adjustedBase + 2;
      
      let label, diffClass;
      if (points < easyThreshold) {
        label = 'Easy';
        diffClass = 'easy';
      } else if (points <= normalMax) {
        label = 'Normal';
        diffClass = 'normal';
      } else if (points <= hardMax) {
        label = 'Hard';
        diffClass = 'hard';
      } else {
        label = 'Deadly';
        diffClass = 'deadly';
      }
      
      let breakdown = '<strong>Base:</strong> (' + BATTLE_CONFIG.BASE_MULTIPLIER + ' × ' + playerCount + ') + ' + BATTLE_CONFIG.BASE_ADDITION + ' = ' + originalBase + '<br>';
      
      if (adjustments.length > 0) {
        breakdown += '<strong>Adjustments:</strong> ' + adjustments.join(', ') + '<br>';
        breakdown += '<strong>Target:</strong> ' + adjustedBase + ' battle points<br>';
      }
      
      breakdown += '<strong>Thresholds:</strong><br>';
      breakdown += 'Easy: ≤' + easyThreshold + ' | Normal: ' + (adjustedBase - 1) + '-' + normalMax + ' | ';
      breakdown += 'Hard: ' + (adjustedBase + 1) + '-' + hardMax + ' | Deadly: ' + (adjustedBase + 3) + '+';
      
      return {
        label: label,
        class: diffClass,
        breakdown: breakdown
      };
    }

    function changeCount(name, delta) {
      if (selectedAdversaries[name]) {
        selectedAdversaries[name] = Math.max(1, selectedAdversaries[name] + delta);
        updateSelectedDisplay();
        updateBattlePoints();
        
        // Update just the badge for this item
        const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
        resultItems.forEach(item => {
          if (item.dataset.advName === name) {
            VDOM.updateResultBadges(item, adversariesData[name]);
          }
        });
      }
    }
    
    function removeSelected(name) {
      delete selectedAdversaries[name];
      updateSelectedDisplay();
      updateBattlePoints();
      
      // Update the specific result item
      const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
      resultItems.forEach(item => {
        if (item.dataset.advName === name) {
          item.classList.remove('selected');
          VDOM.updateResultBadges(item, adversariesData[name]);
        }
      });
    }
    
    function buildEncounterFromSelected() {
      const selectedNames = Object.keys(selectedAdversaries);
      if (selectedNames.length === 0) {
        alert('No adversaries selected!');
        return;
      }
      
      const selections = selectedNames.map(name => ({
        name: name,
        count: selectedAdversaries[name]
      }));
      
      const isHighDamage = DOM_CACHE.get('high-damage').checked;
      const selectedEnvironment = DOM_CACHE.get('environment-select').value;
      
      const btn = DOM_CACHE.get('build-btn');
      btn.innerHTML = 'Creating...';
      btn.disabled = true;
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Encounter created successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error: ' + error.message);
          btn.innerHTML = 'Build Encounter';
          btn.disabled = false;
        })
        .createEncounterFromBuilder(selections, isHighDamage, selectedEnvironment);
    }
    
    function clearFilters() {
      DOM_CACHE.get('tier-filter').value = '';
      DOM_CACHE.get('type-filter').value = '';
      DOM_CACHE.get('range-filter').value = '';
      DOM_CACHE.get('difficulty-min').value = '';
      DOM_CACHE.get('difficulty-max').value = '';
      DOM_CACHE.get('damage-type-filter').value = '';
      DOM_CACHE.get('resistance-filter').value = '';
      DOM_CACHE.get('text-search').value = '';
      DOM_CACHE.get('include-core').checked = true;
      DOM_CACHE.get('include-custom').checked = true;
      SearchManager.lastQuery = ''; // Reset query cache
      initializeAllAdversaries();
    }
    
    initializeDialog();
  `;
}

function createEncounterFromBuilder(selections, isHighDamage = false, selectedEnvironment = '') {
  loadConfiguration();
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const nameResult = ui.prompt(
      'Create Encounter',
      'Enter a name for this encounter:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResult.getSelectedButton() !== ui.Button.OK || !nameResult.getResponseText().trim()) {
      throw new Error('Encounter creation cancelled');
    }
    
    const encounterName = nameResult.getResponseText().trim();
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName(encounterName);
    
    if (existingSheet) {
      const confirmResult = ui.alert(
        'Overwrite Existing Encounter',
        `An encounter named "${encounterName}" already exists. This will completely replace the existing encounter.\n\nDo you want to continue?`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResult !== ui.Button.YES) {
        throw new Error('Encounter creation cancelled');
      }
    }
    
    const allTemplates = getAllAdversaryTemplates();
    const encounterData = [];
    
    for (const selection of selections) {
      const template = allTemplates[selection.name];
      if (template) {
        encounterData.push({
          name: selection.name,
          count: selection.count,
          template: template
        });
      }
    }
    
    if (encounterData.length === 0) {
      throw new Error('No valid adversaries found');
    }
    
    const encounterSheet = getOrCreateSheet(encounterName);
    
    buildEnhancedEncounterLayout(encounterSheet, encounterData, encounterName, isHighDamage, selectedEnvironment);
    
    return { success: true };
    
  } catch (error) {
    console.error('Error creating encounter:', error);
    throw error;
  }
}

function clearDataCache() {
  DATA_CACHE.clearAll();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cache cleared successfully. Data will be reloaded on next use.');
}

function showEncounterBuilder() {
  loadConfiguration();
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = createEncounterBuilderDialog();
  const dialogSize = DIALOG_PRESETS.ENCOUNTER_BUILDER || { WIDTH: 1100, HEIGHT: 750 };
  htmlOutput.setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
  ui.showModalDialog(htmlOutput, 'Encounter Builder');
}

function createEncounterBuilderDialog() {
  const html = generateOptimizedDialogHTML();
  return HtmlService.createHtmlOutput(html);
}

function getDialogData() {
  // Load configuration if needed (will use cache if valid)
  loadConfiguration();
  
  // Check if all data is cached and valid
  if (DATA_CACHE.isDataValid() && DATA_CACHE.adversaries && DATA_CACHE.environments && DATA_CACHE.filterOptions) {
    return {
      adversaries: DATA_CACHE.adversaries,
      environments: DATA_CACHE.environments,
      filterOptions: DATA_CACHE.filterOptions,
      battleValues: BATTLE_VALUES,
      damageTypes: DATA_CACHE.filterOptions.damageTypes || ['physical', 'magic'] // Use extracted or default to both
    };
  }
  
  // Load data (will use cache if available)
  const adversaries = getAllAdversaryTemplates();
  const environments = getAllEnvironmentTemplates();
  
  // Only recalculate filter options if needed
  if (!DATA_CACHE.filterOptions || !DATA_CACHE.isDataValid()) {
    DATA_CACHE.filterOptions = extractFilterOptions(adversaries);
  }
  
  return {
    adversaries: adversaries,
    environments: environments,
    filterOptions: DATA_CACHE.filterOptions,
    battleValues: BATTLE_VALUES,
    damageTypes: DATA_CACHE.filterOptions.damageTypes || ['physical', 'magic'] // Use extracted or default to both
  };
}

function extractFilterOptions(adversaries) {
  const values = Object.values(adversaries);
  
  const tiers = new Set();
  const types = new Set();
  const ranges = new Set();
  const resistances = new Set();
  
  values.forEach(adv => {
    tiers.add(adv.tier);
    types.add(adv.type);
    ranges.add(adv.standardAttack.range);
    
    if (adv.features) {
      adv.features.forEach(feature => {
        const desc = feature.description.toLowerCase();
        if (desc.includes('resistant to') || desc.includes('resistance to')) {
          const match = desc.match(/resistant? to (.+?)(?:\.|$|,)/i);
          if (match) {
            const damageTypes = match[1].split(/\s+and\s+|\s*,\s*|\s+or\s+/);
            damageTypes.forEach(type => {
              const cleanType = type.trim().replace(/\s+damage$/, '');
              if (cleanType && cleanType.length > 1) {
                resistances.add(cleanType);
              }
            });
          }
        }
      });
    }
  });
  
  return {
    tiers: Array.from(tiers).sort(),
    types: Array.from(types).sort(),
    ranges: Array.from(ranges).sort(),
    resistances: Array.from(resistances).sort()
  };
}

function generateOptimizedDialogHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          ${getOptimizedDialogCSS()}
        </style>
      </head>
      <body>
        <div class="loading-overlay" id="loading-overlay">
          <div class="loading-spinner"></div>
          <div class="loading-text">Loading Encounter Builder...</div>
        </div>
        <div class="main-container" id="main-container" style="display: none;">
          <div class="left-panel" id="left-panel"></div>
          <div class="right-panel" id="right-panel"></div>
        </div>
        <script>
          ${generateOptimizedDialogScript()}
        </script>
      </body>
    </html>
  `;
}

function getOptimizedDialogCSS() {
  const maxHeight = SEARCH_CONFIG.RESULTS_MAX_HEIGHT || 450;
  
  return `
    * { box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 15px; 
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      min-height: 100vh;
    }
    .loading-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.95);
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    }
    .loading-spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #2196F3;
      border-radius: 50%;
      width: 50px;
      height: 50px;
      animation: spin 1s linear infinite;
      margin-bottom: 20px;
    }
    .loading-text {
      font-size: 18px;
      color: #333;
      font-weight: 600;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .main-container {
      display: flex;
      gap: 20px;
      max-width: 1400px;
      margin: 0 auto;
      height: calc(100vh - 30px);
    }
    .left-panel {
      flex: 2.2;
      background: white;
      padding: 25px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      height: 100%;
      overflow-y: auto;
    }
    .right-panel {
      flex: 1;
      background: white;
      padding: 25px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      height: 100%;
      overflow-y: auto;
      position: sticky;
      top: 15px;
    }
    .filter-section {
      background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
      padding: 25px;
      border-radius: 12px;
      margin-bottom: 25px;
      border: 2px solid #e3f2fd;
      box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    }
    .filter-header h3 {
      margin: 0 0 20px 0;
      color: #1976d2;
      font-size: 18px;
      font-weight: 600;
    }
    .filter-grid {
      display: flex;
      flex-direction: column;
      gap: 15px;
      margin-bottom: 20px;
    }
    .filter-row {
      display: flex;
      gap: 20px;
    }
    .filter-item {
      display: flex;
      flex-direction: column;
      gap: 8px;
      flex: 1;
    }
    .filter-label {
      font-weight: 600;
      color: #333;
      font-size: 14px;
    }
    select, input {
      padding: 12px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: all 0.3s ease;
      background: white;
    }
    select:focus, input:focus {
      border-color: #1976d2;
      outline: none;
      box-shadow: 0 0 0 3px rgba(25,118,210,0.2);
    }
    .difficulty-range {
      display: flex;
      gap: 12px;
      align-items: center;
    }
    .difficulty-range input { flex: 1; }
    .search-btn, .clear-btn {
      padding: 12px 28px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    .search-btn {
      background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
      color: white;
    }
    .clear-btn {
      background: linear-gradient(135deg, #757575 0%, #616161 100%);
      color: white;
    }
    .results-section {
      background: white;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      max-height: ${maxHeight}px;
      overflow-y: auto;
      overflow-x: hidden;
      word-wrap: break-word;
    }
    .results-count {
      text-align: center;
      padding: 20px;
      font-weight: 600;
      color: #555;
      background: #fafafa;
      position: sticky;
      top: 0;
      z-index: 10;
    }
    .result-item {
      padding: 15px;
      border-bottom: 1px solid #f0f0f0;
      cursor: pointer;
      transition: all 0.3s ease;
      position: relative;
    }
    .result-item:hover {
      background: #f8f9fa;
      transform: translateX(4px);
    }
    .result-item.selected {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border-left: 5px solid #4caf50;
      transform: translateX(8px);
    }
    .result-name {
      font-weight: 700;
      font-size: 16px;
      color: #1976d2;
      margin-bottom: 5px;
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 8px;
    }
    .result-badges {
      display: flex;
      align-items: center;
      gap: 4px;
      flex-shrink: 0;
    }
    .result-stats {
      font-size: 13px;
      color: #666;
      margin-bottom: 8px;
      line-height: 1.4;
    }
    .result-description {
      font-size: 12px;
      color: #444;
      line-height: 1.5;
      font-style: italic;
    }
    .selection-badge {
      background: #4caf50;
      color: white;
      padding: 4px 8px;
      border-radius: 12px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .tier-badge {
      background: #ff9800;
      color: white;
      padding: 2px 6px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .type-badge {
      background: #9c27b0;
      color: white;
      padding: 2px 6px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: bold;
      white-space: nowrap;
    }
    .no-results {
      text-align: center;
      padding: 60px 40px;
      color: #999;
      font-style: italic;
      font-size: 16px;
    }
    .encounter-builder {
      height: 100%;
      display: flex;
      flex-direction: column;
    }
    .builder-header {
      text-align: center;
      margin-bottom: 25px;
      padding-bottom: 20px;
      border-bottom: 3px solid #e3f2fd;
    }
    .builder-header h3 {
      margin: 0;
      color: #1976d2;
      font-size: 22px;
      font-weight: 700;
    }
    .config-section {
      background: #f5f5f5;
      border: 2px solid #bdbdbd;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .config-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 15px;
    }
    .config-row:last-child { margin-bottom: 0; }
    .player-count-input, .player-tier-select {
      width: 80px;
      padding: 8px;
      font-size: 14px;
      text-align: center;
      border: 2px solid #1976d2;
      border-radius: 6px;
    }
    .high-damage-checkbox {
      transform: scale(1.3);
      accent-color: #1976d2;
    }
    .battle-points {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border: 3px solid #4caf50;
      border-radius: 16px;
      padding: 20px;
      margin-bottom: 20px;
      text-align: center;
    }
    .points-display {
      font-size: 24px;
      font-weight: 700;
      color: #2e7d32;
      margin-bottom: 12px;
    }
    .points-breakdown {
      font-size: 13px;
      color: #555;
      line-height: 1.4;
      margin-bottom: 10px;
    }
    .difficulty-breakdown {
      font-size: 11px;
      color: #666;
      margin-top: 6px;
      font-style: italic;
      line-height: 1.3;
    }
    .difficulty-indicator {
      margin-top: 12px;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: 700;
      display: inline-block;
      font-size: 14px;
    }
    .easy { background: #c8e6c9; color: #1b5e20; }
    .normal { background: #fff3c4; color: #000000; }
    .hard { background: #ffcdd2; color: #b71c1c; }
    .deadly { background: #f3e5f5; color: #4a148c; }
    .selected-section {
      background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
      border: 3px solid #4caf50;
      border-radius: 16px;
      margin-bottom: 20px;
      padding: 20px;
      flex: 1;
      display: flex;
      flex-direction: column;
    }
    .selected-header {
      font-weight: 700;
      color: #2e7d32;
      margin-bottom: 15px;
      font-size: 18px;
      text-align: center;
    }
    .selected-list {
      flex: 1;
      overflow-y: auto;
      margin-bottom: 15px;
    }
    .selected-adversary {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px;
      background: white;
      border: 2px solid #a5d6a7;
      border-radius: 10px;
      margin-bottom: 10px;
      font-size: 13px;
    }
    .selected-name {
      font-weight: 700;
      color: #2e7d32;
      flex: 1;
    }
    .selected-stats {
      font-size: 11px;
      color: #666;
      margin-top: 2px;
    }
    .selected-count {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .count-btn {
      background: #4caf50;
      color: white;
      border: none;
      width: 24px;
      height: 24px;
      border-radius: 50%;
      cursor: pointer;
      font-weight: bold;
      font-size: 14px;
    }
    .count-display {
      min-width: 20px;
      text-align: center;
      font-weight: bold;
      font-size: 16px;
      color: #2e7d32;
    }
    .remove-btn {
      background: #f44336;
      color: white;
      border: none;
      padding: 4px 8px;
      border-radius: 6px;
      cursor: pointer;
      font-size: 12px;
      font-weight: bold;
    }
    .build-encounter-btn {
      background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
      color: white;
      border: none;
      padding: 16px 24px;
      border-radius: 12px;
      font-size: 16px;
      font-weight: 700;
      cursor: pointer;
      width: 100%;
      transition: all 0.3s ease;
    }
    .build-encounter-btn:disabled {
      background: #bdbdbd;
      cursor: not-allowed;
    }
    .version-info {
      text-align: center;
      font-size: 11px;
      color: #999;
      margin-top: 15px;
      font-style: italic;
    }
    @media (max-width: 1200px) {
      .main-container { flex-direction: column; height: auto; }
      .right-panel { position: static; margin-top: 20px; }
      .filter-grid { grid-template-columns: 1fr; }
    }
  `;
}

function generateOptimizedDialogScript() {
  // Lazy load configuration values and pass them to client script
  const config = getBattleCalculation();
  const defaultPartySize = config.DEFAULT_PARTY_SIZE || 4;
  const defaultPartyTier = config.DEFAULT_PARTY_TIER || 2;
  
  return `
    // ============================================
    // Enhanced DOM Management System
    // ============================================
    
    let adversariesData = {};
    let environmentsData = {};
    let battleValues = {};
    let damageTypes = [];
    let filterOptions = {};
    let selectedAdversaries = {};
    let lastSearchResults = [];
    let dataLoaded = false;
    
    // Battle calculation configuration - loaded from server
    const BATTLE_CONFIG = {
      BASE_MULTIPLIER: ${config.BASE_MULTIPLIER || 3},
      BASE_ADDITION: ${config.BASE_ADDITION || 2},
      MULTIPLE_SOLOS_ADJUSTMENT: ${config.MULTIPLE_SOLOS_ADJUSTMENT || -2},
      LOWER_TIER_BONUS: ${config.LOWER_TIER_BONUS || 1},
      NO_ELITES_BONUS: ${config.NO_ELITES_BONUS || 1},
      HIGH_DAMAGE_PENALTY: ${config.HIGH_DAMAGE_PENALTY || -2},
      MIN_SOLOS_FOR_ADJUSTMENT: ${config.MIN_SOLOS_FOR_ADJUSTMENT || 2},
      DEFAULT_PARTY_SIZE: ${defaultPartySize},
      DEFAULT_PARTY_TIER: ${defaultPartyTier}
    };
    
    // DOM element cache to avoid repeated queries
    const DOM_CACHE = {
      elements: new Map(),
      
      get(id) {
        if (!this.elements.has(id)) {
          this.elements.set(id, document.getElementById(id));
        }
        return this.elements.get(id);
      },
      
      clear() {
        this.elements.clear();
      }
    };
    
    // Virtual DOM for search results
    const VDOM = {
      currentResults: new Map(),
      resultContainer: null,
      resultPool: [], // Reusable DOM elements
      maxPoolSize: 100,
      
      init() {
        this.resultContainer = DOM_CACHE.get('results-section');
      },
      
      // Get or create a result item element
      getResultElement() {
        if (this.resultPool.length > 0) {
          return this.resultPool.pop();
        }
        
        const div = document.createElement('div');
        div.className = 'result-item';
        return div;
      },
      
      // Return element to pool for reuse
      releaseElement(element) {
        if (this.resultPool.length < this.maxPoolSize) {
          element.className = 'result-item';
          element.onclick = null;
          this.resultPool.push(element);
        }
      },
      
      // Efficiently update only changed results
      updateResults(results) {
        const newResultsMap = new Map();
        const fragment = document.createDocumentFragment();
        let hasChanges = false;
        
        // Check if we need full rebuild (different result count or first render)
        if (this.currentResults.size === 0 || this.currentResults.size !== results.length) {
          hasChanges = true;
          this.renderFullResults(results);
          return;
        }
        
        // Incremental update for same-size result sets
        const container = this.resultContainer;
        const existingNodes = container.querySelectorAll('.result-item');
        
        results.forEach((adv, index) => {
          const existingNode = existingNodes[index];
          if (existingNode) {
            const oldName = existingNode.dataset.advName;
            if (oldName !== adv.name) {
              // Different adversary, update the node
              this.updateResultNode(existingNode, adv);
              hasChanges = true;
            } else {
              // Same adversary, just update selection state
              const isSelected = adv.name in selectedAdversaries && selectedAdversaries[adv.name] > 0;
              const wasSelected = existingNode.classList.contains('selected');
              if (isSelected !== wasSelected) {
                existingNode.classList.toggle('selected', isSelected);
                this.updateResultBadges(existingNode, adv);
                hasChanges = true;
              }
            }
          }
          newResultsMap.set(adv.name, adv);
        });
        
        this.currentResults = newResultsMap;
      },
      
      // Full render when structure changes significantly
      renderFullResults(results) {
        const container = this.resultContainer;
        
        // Return all current elements to pool
        const existingItems = container.querySelectorAll('.result-item');
        existingItems.forEach(item => this.releaseElement(item));
        
        // Use DocumentFragment for batch DOM operations
        const fragment = document.createDocumentFragment();
        
        // Add count header
        const countDiv = document.createElement('div');
        countDiv.className = 'results-count';
        countDiv.textContent = results.length + ' adversar' + (results.length === 1 ? 'y' : 'ies') + ' found';
        fragment.appendChild(countDiv);
        
        // Add result items
        results.forEach(adv => {
          const div = this.getResultElement();
          this.updateResultNode(div, adv);
          fragment.appendChild(div);
        });
        
        // Single DOM update
        container.innerHTML = '';
        container.appendChild(fragment);
        
        // Update cache
        this.currentResults.clear();
        results.forEach(adv => this.currentResults.set(adv.name, adv));
      },
      
      // Update a single result node
      updateResultNode(node, adv) {
        node.dataset.advName = adv.name;
        node.className = 'result-item' + (selectedAdversaries[adv.name] ? ' selected' : '');
        node.onclick = () => toggleAdversarySelection(adv.name);
        
        // Use innerHTML once for better performance
        node.innerHTML = this.getResultHTML(adv);
      },
      
      // Update only badges without rebuilding entire node
      updateResultBadges(node, adv) {
        const badgesContainer = node.querySelector('.result-badges');
        if (badgesContainer) {
          badgesContainer.innerHTML = this.getBadgesHTML(adv);
        }
      },
      
      // Generate HTML strings (faster than DOM manipulation)
      getResultHTML(adv) {
        return \`
          <div class="result-name">
            <span>\${adv.name}</span>
            <div class="result-badges">\${this.getBadgesHTML(adv)}</div>
          </div>
          <div class="result-stats">\${this.getStatsHTML(adv)}</div>
          <div class="result-description">\${adv.description}</div>
        \`;
      },
      
      getBadgesHTML(adv) {
        let html = \`<span class="tier-badge">T\${adv.tier}</span>\`;
        html += \`<span class="type-badge">\${adv.type}</span>\`;
        // FIX: Proper check for selection
        if (adv.name in selectedAdversaries && selectedAdversaries[adv.name] > 0) {
          html += \`<span class="selection-badge">×\${selectedAdversaries[adv.name]}</span>\`;
        }
        return html;
      },
      
      getStatsHTML(adv) {
        return \`<strong>Difficulty:</strong> \${adv.difficulty} | <strong>HP:</strong> \${adv.hp} | <strong>Stress:</strong> \${adv.stress} | <strong>Value:</strong> \${getBattleValueJS(adv.type)}<br>
                <strong>Attack:</strong> +\${adv.attackModifier} \${adv.standardAttack.name} (\${adv.standardAttack.range}, \${adv.standardAttack.damage})\`;
      }
    };
    
    // Debounced search to reduce unnecessary updates
    const SearchManager = {
      searchTimeout: null,
      lastQuery: '',
      
      debounceSearch() {
        if (this.searchTimeout) {
          clearTimeout(this.searchTimeout);
        }
        
        this.searchTimeout = setTimeout(() => {
          this.executeSearch();
        }, 150); // Reduced from 300ms for snappier feel
      },
      
      executeSearch() {
        if (!dataLoaded) return;
        
        const filters = this.getFilterValues();
        const queryKey = JSON.stringify(filters);
        
        // Skip if query hasn't changed
        if (queryKey === this.lastQuery) return;
        this.lastQuery = queryKey;
        
        let results = Object.values(adversariesData);
        results = this.applyFilters(results, filters);
        results.sort((a, b) => a.name.localeCompare(b.name));
        
        lastSearchResults = results;
        displayResults(results);
      },
      
      getFilterValues() {
        return {
          tier: DOM_CACHE.get('tier-filter').value,
          type: DOM_CACHE.get('type-filter').value,
          range: DOM_CACHE.get('range-filter').value,
          difficultyMin: DOM_CACHE.get('difficulty-min').value,
          difficultyMax: DOM_CACHE.get('difficulty-max').value,
          damageType: DOM_CACHE.get('damage-type-filter').value,
          resistance: DOM_CACHE.get('resistance-filter').value,
          textSearch: DOM_CACHE.get('text-search').value.toLowerCase(),
          includeCore: DOM_CACHE.get('include-core').checked,
          includeCustom: DOM_CACHE.get('include-custom').checked
        };
      },
      
      applyFilters(results, filters) {
        // Source filter first (most restrictive)
        if (!filters.includeCore || !filters.includeCustom) {
          results = results.filter(adv => {
            const isCustom = adv.isCustom === true;
            return (filters.includeCore || isCustom) && (filters.includeCustom || !isCustom);
          });
        }
        
        // Apply other filters only if set
        if (filters.tier) {
          results = results.filter(adv => adv.tier == filters.tier);
        }
        if (filters.type) {
          results = results.filter(adv => adv.type === filters.type);
        }
        if (filters.range) {
          results = results.filter(adv => adv.standardAttack.range === filters.range);
        }
        if (filters.difficultyMin) {
          const min = parseInt(filters.difficultyMin);
          results = results.filter(adv => adv.difficulty >= min);
        }
        if (filters.difficultyMax) {
          const max = parseInt(filters.difficultyMax);
          results = results.filter(adv => adv.difficulty <= max);
        }
        
        // Text-based filters (more expensive)
        if (filters.damageType) {
          results = results.filter(adv => {
            if (!adv.standardAttack || !adv.standardAttack.damage) return false;
            const damageStr = adv.standardAttack.damage.toLowerCase().trim();
            
            if (filters.damageType === 'physical') {
              // If it contains "physical" or "phy", it's physical
              // If it has NO damage type specified (just dice), assume physical per Daggerheart rules
              const hasPhysical = damageStr.includes('physical') || damageStr.includes('phy');
              const hasMagic = damageStr.includes('magic') || damageStr.includes('mag');
              
              return hasPhysical || (!hasPhysical && !hasMagic);
            } else if (filters.damageType === 'magic') {
              // Only include if explicitly says magic/mag
              return damageStr.includes('magic') || damageStr.includes('mag');
            }
            return false;
          });
        }
        
        if (filters.resistance) {
          results = results.filter(adv => {
            if (!adv.features) return false;
            return adv.features.some(f => {
              const desc = f.description.toLowerCase();
              return desc.includes('resistant to ' + filters.resistance) || 
                     desc.includes('resistance to ' + filters.resistance);
            });
          });
        }
        
        // Full text search (most expensive, do last)
        if (filters.textSearch) {
          results = results.filter(adv => {
            // Cache searchable text on adversary object
            if (!adv._searchCache) {
              adv._searchCache = [
                adv.name, adv.description, adv.motives,
                adv.standardAttack.name, adv.standardAttack.damage,
                ...adv.experience,
                ...adv.features.map(f => f.name + ' ' + f.description)
              ].join(' ').toLowerCase();
            }
            return adv._searchCache.includes(filters.textSearch);
          });
        }
        
        return results;
      }
    };
    
    // Optimized initialization
    function initializeDialog() {
      google.script.run
        .withSuccessHandler(onDataLoaded)
        .withFailureHandler(onDataLoadError)
        .getDialogData();
    }

    function onDataLoaded(data) {
      adversariesData = data.adversaries;
      environmentsData = data.environments;
      battleValues = data.battleValues;
      damageTypes = data.damageTypes;
      filterOptions = data.filterOptions;
      dataLoaded = true;
      
      buildUI();
      VDOM.init();
      
      DOM_CACHE.get('loading-overlay').style.display = 'none';
      DOM_CACHE.get('main-container').style.display = 'flex';
      
      // Use requestAnimationFrame for smooth initial render
      requestAnimationFrame(() => {
        initializeAllAdversaries();
        updateBattlePoints();
      });
    }

    function onDataLoadError(error) {
      console.error('Error loading data:', error);
      DOM_CACHE.get('loading-overlay').innerHTML = 
        '<div style="color: red;">Error loading data. Please try again.</div>';
    }

    function buildUI() {
      const leftPanel = DOM_CACHE.get('left-panel');
      const rightPanel = DOM_CACHE.get('right-panel');
      
      leftPanel.innerHTML = generateLeftPanelHTML();
      rightPanel.innerHTML = generateRightPanelHTML();
      
      // Cache frequently used elements
      DOM_CACHE.clear();
    }

    function generateLeftPanelHTML() {
      const html = [];
      html.push('<div class="filter-section">');
      html.push('<div class="filter-header"><h3>Adversary Filters</h3></div>');
      html.push('<div class="filter-grid">');
      
      html.push('<div class="filter-row">');
      html.push(createFilterItem('Tier', 'tier-filter', 'Any Tier', 
        filterOptions.tiers.map(t => ({value: t, label: 'Tier ' + t}))));
      html.push(createFilterItem('Type', 'type-filter', 'Any Type',
        filterOptions.types.map(t => ({value: t, label: t}))));
      html.push(createDifficultyRangeFilter());
      html.push('</div>');
      
      html.push('<div class="filter-row">');
      html.push(createFilterItem('Range', 'range-filter', 'Any Range',
        filterOptions.ranges.map(r => ({value: r, label: r}))));
      html.push(createFilterItem('Damage Type', 'damage-type-filter', 'Any Damage Type',
        damageTypes.map(t => ({value: t, label: t.charAt(0).toUpperCase() + t.slice(1)}))));
      html.push(createFilterItem('Damage Resistance', 'resistance-filter', 'Any Resistances',
        filterOptions.resistances.map(r => ({value: r, label: r.charAt(0).toUpperCase() + r.slice(1)}))));
      html.push('</div>');
      
      html.push('<div class="filter-row">');
      html.push(createTextSearchFilter());
      html.push(createSourceFilter());
      html.push(createClearButton());
      html.push('</div>');
      
      html.push('</div></div>');
      html.push('<div class="results-section" id="results-section">');
      html.push('<div class="no-results">Loading adversaries...</div>');
      html.push('</div>');
      
      return html.join('');
    }

    function createFilterItem(label, id, defaultText, options) {
      const html = [];
      html.push('<div class="filter-item">');
      html.push('<label class="filter-label">' + label + '</label>');
      html.push('<select id="' + id + '" onchange="searchAdversaries()">');
      html.push('<option value="">' + defaultText + '</option>');
      options.forEach(opt => {
        html.push('<option value="' + opt.value + '">' + opt.label + '</option>');
      });
      html.push('</select></div>');
      return html.join('');
    }

    function createDifficultyRangeFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Difficulty Range</label>' +
        '<div class="difficulty-range">' +
        '<input type="number" id="difficulty-min" placeholder="Min" min="1" max="30" onchange="searchAdversaries()">' +
        '<span>to</span>' +
        '<input type="number" id="difficulty-max" placeholder="Max" min="1" max="30" onchange="searchAdversaries()">' +
        '</div></div>';
    }

    function createTextSearchFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Text Search</label>' +
        '<input type="text" id="text-search" placeholder="Search names, descriptions, features..." oninput="searchAdversaries()">' +
        '</div>';
    }

    function createSourceFilter() {
      return '<div class="filter-item">' +
        '<label class="filter-label">Adversary Sources</label>' +
        '<div style="display: flex; gap: 15px; align-items: center;">' +
        '<label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">' +
        '<input type="checkbox" id="include-core" checked onchange="searchAdversaries()" style="transform: scale(1.2);">' +
        'Core</label>' +
        '<label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">' +
        '<input type="checkbox" id="include-custom" checked onchange="searchAdversaries()" style="transform: scale(1.2);">' +
        'Custom</label></div></div>';
    }

    function createClearButton() {
      return '<div class="filter-item" style="display: flex; flex-direction: column;">' +
        '<label class="filter-label" style="visibility: hidden;">Clear</label>' +
        '<button class="clear-btn" onclick="clearFilters()">Clear Filters</button>' +
        '</div>';
    }

    function generateRightPanelHTML() {
      const html = [];
      html.push('<div class="encounter-builder">');
      html.push('<div class="builder-header"><h3>Encounter Builder</h3></div>');
      
      html.push('<div class="config-section">');
      html.push(createConfigRow('Players:', 'player-count', 'number', '${defaultPartySize}', 'player-count-input'));
      html.push(createTierSelect());
      html.push(createHighDamageCheckbox());
      html.push(createEnvironmentSelect());
      html.push('</div>');
      
      html.push('<div class="battle-points">');
      html.push('<div class="points-display" id="total-points">0 Battle Points</div>');
      html.push('<div class="points-breakdown" id="points-breakdown">Select adversaries to see difficulty</div>');
      html.push('<div class="difficulty-breakdown" id="difficulty-breakdown">Configure encounter to see calculation details</div>');
      html.push('<div class="difficulty-indicator" id="difficulty-indicator" style="display: none;"></div>');
      html.push('</div>');
      
      html.push('<div class="selected-section" id="selected-section" style="display: none;">');
      html.push('<div class="selected-header">Selected Adversaries</div>');
      html.push('<div class="selected-list" id="selected-list"></div>');
      html.push('<button class="build-encounter-btn" id="build-btn" onclick="buildEncounterFromSelected()">Build Encounter</button>');
      html.push('</div>');
      
      html.push('<div class="version-info">Daggerheart Combat System v0.2</div>');
      html.push('</div>');
      
      return html.join('');
    }

    function createConfigRow(label, id, type, value, className) {
      return '<div class="config-row">' +
        '<label><strong>' + label + '</strong></label>' +
        '<input type="' + type + '" id="' + id + '" class="' + className + '" value="' + value + '" min="1" max="8" onchange="updateBattlePoints()">' +
        '</div>';
    }

    function createTierSelect() {
      const defaultTier = ${defaultPartyTier};
      return '<div class="config-row">' +
        '<label><strong>Tier:</strong></label>' +
        '<select id="player-tier" class="player-tier-select" onchange="updateBattlePoints()">' +
        '<option value="1"' + (defaultTier === 1 ? ' selected' : '') + '>Tier 1</option>' +
        '<option value="2"' + (defaultTier === 2 ? ' selected' : '') + '>Tier 2</option>' +
        '<option value="3"' + (defaultTier === 3 ? ' selected' : '') + '>Tier 3</option>' +
        '<option value="4"' + (defaultTier === 4 ? ' selected' : '') + '>Tier 4</option>' +
        '</select></div>';
    }

    function createHighDamageCheckbox() {
      return '<div class="config-row">' +
        '<label><strong>High Damage:</strong></label>' +
        '<input type="checkbox" id="high-damage" class="high-damage-checkbox" onchange="updateBattlePoints()">' +
        '</div>';
    }

    function createEnvironmentSelect() {
      const envNames = Object.keys(environmentsData).sort();
      const options = envNames.map(name => '<option value="' + name + '">' + name + '</option>').join('');
      return '<div class="config-row">' +
        '<label><strong>Environment:</strong></label>' +
        '<select id="environment-select" class="player-tier-select" style="width: 200px;">' +
        '<option value="">No Environment</option>' + options +
        '</select></div>';
    }

    function getBattleValueJS(adversaryType) {
      return battleValues[adversaryType] || 2;
    }

    function initializeAllAdversaries() {
      if (!dataLoaded) return;
      let results = Object.values(adversariesData);
      results.sort((a, b) => a.name.localeCompare(b.name));
      lastSearchResults = results;
      displayResults(results);
    }

    // Use new optimized search
    function searchAdversaries() {
      SearchManager.debounceSearch();
    }

    // Use new optimized display
    function displayResults(results) {
      if (results.length === 0) {
        VDOM.resultContainer.innerHTML = '<div class="no-results">No adversaries found matching your criteria</div>';
        VDOM.currentResults.clear();
        return;
      }
      
      VDOM.updateResults(results);
    }
    
    function toggleAdversarySelection(name) {
      if (selectedAdversaries[name]) {
        delete selectedAdversaries[name];
      } else {
        selectedAdversaries[name] = 1;
      }
      updateSelectedDisplay();
      updateBattlePoints();
      
      // Optimized: only update the specific result item
      const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
      resultItems.forEach(item => {
        if (item.dataset.advName === name) {
          item.classList.toggle('selected', selectedAdversaries[name]);
          VDOM.updateResultBadges(item, adversariesData[name]);
        }
      });
    }
    
    // Optimized selected display with fragment
    function updateSelectedDisplay() {
      const section = DOM_CACHE.get('selected-section');
      const list = DOM_CACHE.get('selected-list');
      
      const names = Object.keys(selectedAdversaries);
      
      if (names.length === 0) {
        section.style.display = 'none';
        return;
      }
      
      section.style.display = 'block';
      
      // Use DocumentFragment for batch updates
      const fragment = document.createDocumentFragment();
      
      names.forEach(name => {
        const adv = adversariesData[name];
        const count = selectedAdversaries[name];
        
        const div = document.createElement('div');
        div.className = 'selected-adversary';
        div.innerHTML = createSelectedAdversaryHTML(name, adv, count);
        fragment.appendChild(div);
      });
      
      list.innerHTML = '';
      list.appendChild(fragment);
    }

    function createSelectedAdversaryHTML(name, adv, count) {
      return '<div style="flex: 1;">' +
        '<div class="selected-name">' + name + '</div>' +
        '<div class="selected-stats">T' + adv.tier + ' ' + adv.type + 
        ' | Diff: ' + adv.difficulty + ' | Value: ' + getBattleValueJS(adv.type) + '</div>' +
        '</div>' +
        '<div class="selected-count">' +
        '<button class="count-btn" onclick="changeCount(\\'' + name.replace(/'/g, "\\'") + '\\', -1)">-</button>' +
        '<span class="count-display">' + count + '</span>' +
        '<button class="count-btn" onclick="changeCount(\\'' + name.replace(/'/g, "\\'") + '\\', 1)">+</button>' +
        '<button class="remove-btn" onclick="removeSelected(\\'' + name.replace(/'/g, "\\'") + '\\')">×</button>' +
        '</div>';
    }

    // Cached battle points calculation
    let lastBattlePointsState = '';
    let lastBattlePointsResult = null;
    
    function updateBattlePoints() {
      if (!dataLoaded) return;
      
      try {
        // Create state key to check if recalculation needed
        const stateKey = JSON.stringify({
          selected: selectedAdversaries,
          playerCount: DOM_CACHE.get('player-count').value,
          playerTier: DOM_CACHE.get('player-tier').value,
          highDamage: DOM_CACHE.get('high-damage').checked
        });
        
        // Use cached result if nothing changed
        if (stateKey === lastBattlePointsState && lastBattlePointsResult) {
          applyBattlePointsDisplay(lastBattlePointsResult);
          return;
        }
        
        let totalPoints = 0;
        const breakdown = [];
        const selectedList = [];
        
        const playerCount = parseInt(DOM_CACHE.get('player-count').value) || BATTLE_CONFIG.DEFAULT_PARTY_SIZE;
        
        Object.keys(selectedAdversaries).forEach(name => {
          const count = selectedAdversaries[name];
          if (adversariesData[name] && count > 0) {
            const template = adversariesData[name];
            let points;
            
            if (template.type === 'Minion') {
              const groups = Math.ceil(count / playerCount);
              points = groups;
              breakdown.push(count + 'x ' + name + ' (' + groups + ' group' + 
                            (groups > 1 ? 's' : '') + ' = ' + points + ' pts)');
            } else {
              const value = getBattleValueJS(template.type);
              points = value * count;
              breakdown.push(count + 'x ' + name + ' (' + points + ' pts)');
            }
            
            totalPoints += points;
            
            for (let i = 0; i < count; i++) {
              selectedList.push(template);
            }
          }
        });
        
        const result = {
          totalPoints: totalPoints,
          breakdown: breakdown,
          selectedList: selectedList,
          playerCount: playerCount,
          playerTier: parseInt(DOM_CACHE.get('player-tier').value) || BATTLE_CONFIG.DEFAULT_PARTY_TIER,
          isHighDamage: DOM_CACHE.get('high-damage').checked
        };
        
        // Cache the result
        lastBattlePointsState = stateKey;
        lastBattlePointsResult = result;
        
        applyBattlePointsDisplay(result);
        
      } catch (error) {
        console.error('Error in updateBattlePoints:', error);
      }
    }
    
    function applyBattlePointsDisplay(result) {
      DOM_CACHE.get('total-points').textContent = result.totalPoints + ' Battle Points';
      
      if (result.breakdown.length > 0) {
        DOM_CACHE.get('points-breakdown').innerHTML = result.breakdown.join('<br>');
        
        const difficulty = getDifficultyLevel(
          result.totalPoints, 
          result.playerCount, 
          result.playerTier, 
          result.selectedList, 
          result.isHighDamage
        );
        
        const indicator = DOM_CACHE.get('difficulty-indicator');
        indicator.textContent = difficulty.label;
        indicator.className = 'difficulty-indicator ' + difficulty.class;
        indicator.style.display = 'inline-block';
        
        DOM_CACHE.get('difficulty-breakdown').innerHTML = difficulty.breakdown;
        DOM_CACHE.get('build-btn').disabled = false;
      } else {
        DOM_CACHE.get('points-breakdown').textContent = 'Select adversaries to see difficulty';
        DOM_CACHE.get('difficulty-breakdown').textContent = 'Configure encounter to see calculation details';
        DOM_CACHE.get('difficulty-indicator').style.display = 'none';
        DOM_CACHE.get('build-btn').disabled = true;
      }
    }

    function getDifficultyLevel(points, playerCount, playerTier, adversaries, isHighDamage) {
      const originalBase = (BATTLE_CONFIG.BASE_MULTIPLIER * playerCount) + BATTLE_CONFIG.BASE_ADDITION;
      let adjustedBase = originalBase;
      let adjustments = [];
      
      const soloCount = adversaries.filter(adv => adv.type === 'Solo').length;
      if (soloCount >= BATTLE_CONFIG.MIN_SOLOS_FOR_ADJUSTMENT) {
        adjustedBase += BATTLE_CONFIG.MULTIPLE_SOLOS_ADJUSTMENT;
        adjustments.push(BATTLE_CONFIG.MULTIPLE_SOLOS_ADJUSTMENT + ' (multiple solos)');
      }
      
      const hasLowerTier = adversaries.some(adv => adv.tier < playerTier);
      if (hasLowerTier) {
        adjustedBase += BATTLE_CONFIG.LOWER_TIER_BONUS;
        adjustments.push('+' + BATTLE_CONFIG.LOWER_TIER_BONUS + ' (lower tier adversaries)');
      }
      
      const hasElites = adversaries.some(adv => 
        ['Bruiser', 'Horde', 'Leader', 'Solo'].includes(adv.type)
      );
      if (!hasElites && adversaries.length > 0) {
        adjustedBase += BATTLE_CONFIG.NO_ELITES_BONUS;
        adjustments.push('+' + BATTLE_CONFIG.NO_ELITES_BONUS + ' (no elite types)');
      }
      
      if (isHighDamage) {
        adjustedBase += BATTLE_CONFIG.HIGH_DAMAGE_PENALTY;
        adjustments.push(BATTLE_CONFIG.HIGH_DAMAGE_PENALTY + ' (high damage encounter)');
      }
      
      const easyThreshold = adjustedBase - 1;
      const normalMax = adjustedBase;
      const hardMax = adjustedBase + 2;
      
      let label, diffClass;
      if (points < easyThreshold) {
        label = 'Easy';
        diffClass = 'easy';
      } else if (points <= normalMax) {
        label = 'Normal';
        diffClass = 'normal';
      } else if (points <= hardMax) {
        label = 'Hard';
        diffClass = 'hard';
      } else {
        label = 'Deadly';
        diffClass = 'deadly';
      }
      
      let breakdown = '<strong>Base:</strong> (' + BATTLE_CONFIG.BASE_MULTIPLIER + ' × ' + playerCount + ') + ' + BATTLE_CONFIG.BASE_ADDITION + ' = ' + originalBase + '<br>';
      
      if (adjustments.length > 0) {
        breakdown += '<strong>Adjustments:</strong> ' + adjustments.join(', ') + '<br>';
        breakdown += '<strong>Target:</strong> ' + adjustedBase + ' battle points<br>';
      }
      
      breakdown += '<strong>Thresholds:</strong><br>';
      breakdown += 'Easy: ≤' + easyThreshold + ' | Normal: ' + (adjustedBase - 1) + '-' + normalMax + ' | ';
      breakdown += 'Hard: ' + (adjustedBase + 1) + '-' + hardMax + ' | Deadly: ' + (adjustedBase + 3) + '+';
      
      return {
        label: label,
        class: diffClass,
        breakdown: breakdown
      };
    }

    function changeCount(name, delta) {
      if (selectedAdversaries[name]) {
        selectedAdversaries[name] = Math.max(1, selectedAdversaries[name] + delta);
        updateSelectedDisplay();
        updateBattlePoints();
        
        // Update just the badge for this item
        const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
        resultItems.forEach(item => {
          if (item.dataset.advName === name) {
            VDOM.updateResultBadges(item, adversariesData[name]);
          }
        });
      }
    }
    
    function removeSelected(name) {
      delete selectedAdversaries[name];
      updateSelectedDisplay();
      updateBattlePoints();
      
      // Update the specific result item
      const resultItems = VDOM.resultContainer.querySelectorAll('.result-item');
      resultItems.forEach(item => {
        if (item.dataset.advName === name) {
          item.classList.remove('selected');
          VDOM.updateResultBadges(item, adversariesData[name]);
        }
      });
    }
    
    function buildEncounterFromSelected() {
      const selectedNames = Object.keys(selectedAdversaries);
      if (selectedNames.length === 0) {
        alert('No adversaries selected!');
        return;
      }
      
      const selections = selectedNames.map(name => ({
        name: name,
        count: selectedAdversaries[name]
      }));
      
      const isHighDamage = DOM_CACHE.get('high-damage').checked;
      const selectedEnvironment = DOM_CACHE.get('environment-select').value;
      
      const btn = DOM_CACHE.get('build-btn');
      btn.innerHTML = 'Creating...';
      btn.disabled = true;
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Encounter created successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error: ' + error.message);
          btn.innerHTML = 'Build Encounter';
          btn.disabled = false;
        })
        .createEncounterFromBuilder(selections, isHighDamage, selectedEnvironment);
    }
    
    function clearFilters() {
      DOM_CACHE.get('tier-filter').value = '';
      DOM_CACHE.get('type-filter').value = '';
      DOM_CACHE.get('range-filter').value = '';
      DOM_CACHE.get('difficulty-min').value = '';
      DOM_CACHE.get('difficulty-max').value = '';
      DOM_CACHE.get('damage-type-filter').value = '';
      DOM_CACHE.get('resistance-filter').value = '';
      DOM_CACHE.get('text-search').value = '';
      DOM_CACHE.get('include-core').checked = true;
      DOM_CACHE.get('include-custom').checked = true;
      SearchManager.lastQuery = ''; // Reset query cache
      initializeAllAdversaries();
    }
    
    initializeDialog();
  `;
}

function createEncounterFromBuilder(selections, isHighDamage = false, selectedEnvironment = '') {
  loadConfiguration();
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const nameResult = ui.prompt(
      'Create Encounter',
      'Enter a name for this encounter:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResult.getSelectedButton() !== ui.Button.OK || !nameResult.getResponseText().trim()) {
      throw new Error('Encounter creation cancelled');
    }
    
    const encounterName = nameResult.getResponseText().trim();
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = spreadsheet.getSheetByName(encounterName);
    
    if (existingSheet) {
      const confirmResult = ui.alert(
        'Overwrite Existing Encounter',
        `An encounter named "${encounterName}" already exists. This will completely replace the existing encounter.\n\nDo you want to continue?`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResult !== ui.Button.YES) {
        throw new Error('Encounter creation cancelled');
      }
    }
    
    const allTemplates = getAllAdversaryTemplates();
    const encounterData = [];
    
    for (const selection of selections) {
      const template = allTemplates[selection.name];
      if (template) {
        encounterData.push({
          name: selection.name,
          count: selection.count,
          template: template
        });
      }
    }
    
    if (encounterData.length === 0) {
      throw new Error('No valid adversaries found');
    }
    
    const encounterSheet = getOrCreateSheet(encounterName);
    
    buildEnhancedEncounterLayout(encounterSheet, encounterData, encounterName, isHighDamage, selectedEnvironment);
    
    return { success: true };
    
  } catch (error) {
    console.error('Error creating encounter:', error);
    throw error;
  }
}

function buildEnhancedEncounterLayout(sheet, encounterData, encounterName, isHighDamage, selectedEnvironment) {
  let row = 1;
  
  createEncounterTitle(sheet, row, encounterName);
  row += 1;
  
  if (isHighDamage) {
    createHighDamageWarning(sheet, row);
    row += 1;
  }
  
  createEncounterSummary(sheet, row, encounterData);
  row += 2;
  
  row = processAdversaryBlocks(sheet, row, encounterData);
  
  if (selectedEnvironment) {
    processEnvironmentBlock(sheet, row, selectedEnvironment);
  }
  
  formatEncounterSheet(sheet);
}

function formatEncounterSheet(sheet) {
  // Lazy load sheet formatting only when needed
  const formatting = getSheetFormatting();
  
  sheet.setColumnWidth(1, formatting.NAME_COLUMN_WIDTH);
  sheet.setColumnWidth(2, formatting.DIFFICULTY_COLUMN_WIDTH);
  sheet.setColumnWidth(3, formatting.THRESHOLDS_COLUMN_WIDTH);
  sheet.setColumnWidth(4, formatting.ATTACK_COLUMN_WIDTH);
  sheet.setColumnWidth(5, formatting.STANDARD_ATTACK_COLUMN_WIDTH);
  
  const maxCol = sheet.getLastColumn();
  for (let col = 6; col <= maxCol - 1; col++) {
    sheet.setColumnWidth(col, formatting.CHECKBOX_COLUMN_WIDTH);
  }
  if (maxCol > 5) {
    sheet.setColumnWidth(maxCol, formatting.NOTES_COLUMN_WIDTH);
  }
}
function createEncounterTitle(sheet, row, encounterName) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const titleRange = sheet.getRange(row, 1, 1, maxCols);
  titleRange.merge();
  titleRange.setValue(encounterName.toUpperCase());
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(18);
  titleRange.setHorizontalAlignment('center');
  titleRange.setBackground('#424242');
  titleRange.setFontColor('white');
}

function createHighDamageWarning(sheet, row) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const noteRange = sheet.getRange(row, 1, 1, maxCols);
  noteRange.merge();
  noteRange.setValue('HIGH DAMAGE ENCOUNTER: Add +1d4 to all damage rolls');
  noteRange.setBackground('#ffecb3');
  noteRange.setFontColor('#e65100');
  noteRange.setFontWeight('bold');
  noteRange.setHorizontalAlignment('center');
}

function createEncounterSummary(sheet, row, encounterData) {
  // Lazy load sheet formatting
  const maxCols = getSheetFormatting().MAX_COLUMNS_IN_SHEET || 12;
  const totalEnemies = encounterData.reduce((sum, ed) => sum + ed.count, 0);
  let totalValue = 0;
  
  // Lazy load battle calculation
  const battleCalc = getBattleCalculation();
  
  encounterData.forEach(ed => {
    if (ed.template.type === 'Minion') {
      const defaultPartySize = battleCalc.DEFAULT_PARTY_SIZE || 4;
      totalValue += Math.ceil(ed.count / defaultPartySize);
    } else {
      totalValue += getBattleValue(ed.template.type) * ed.count;
    }
  });
  
  const summaryRange = sheet.getRange(row, 1, 1, maxCols);
  summaryRange.merge();
  summaryRange.setValue(`${totalEnemies} adversaries | ${totalValue} battle points`);
  summaryRange.setBackground('#eeeeee');
  summaryRange.setFontColor('#424242');
  summaryRange.setHorizontalAlignment('center');
}

function processAdversaryBlocks(sheet, startRow, encounterData) {
  let row = startRow;
  
  for (const encounter of encounterData) {
    const { name, count, template } = encounter;
    const totalCols = 5 + template.hp + template.stress + 1;
    
    row = createAdversaryHeader(sheet, row, totalCols, count, name, template);
    row = createAdversaryTable(sheet, row, totalCols, count, name, template);
    row = createAdversaryExperience(sheet, row, totalCols, template);
    row = createAdversaryFeatures(sheet, row, totalCols, template);
    row += 1;
  }
  
  return row;
}

function createAdversaryHeader(sheet, row, totalCols, count, name, template) {
  const countText = count === 1 ? '1x' : `${count}x`;
  const headerRange = sheet.getRange(row, 1, 1, totalCols);
  headerRange.merge();
  headerRange.setValue(`${countText} ${name} (Tier ${template.tier} ${template.type})`);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#757575');
  headerRange.setFontColor('white');
  return row + 1;
}

function createAdversaryTable(sheet, row, totalCols, count, name, template) {
  const headers = ['Name', 'Difficulty', 'Thresholds', 'Attack', 'Standard Attack'];
  
  for (let i = 1; i <= template.hp; i++) {
    headers.push(`❤️${i}`);
  }
  for (let i = 1; i <= template.stress; i++) {
    headers.push(`😰${i}`);
  }
  headers.push('Notes');
  
  const tableHeaderRange = sheet.getRange(row, 1, 1, totalCols);
  tableHeaderRange.setValues([headers]);
  tableHeaderRange.setFontWeight('bold');
  tableHeaderRange.setBackground('#e8f5e8');
  tableHeaderRange.setFontColor('#2e7d32');
  tableHeaderRange.setHorizontalAlignment('center');
  row++;
  
  for (let i = 1; i <= count; i++) {
    const adversaryName = count === 1 ? name : `${name} ${i}`;
    const thresholdText = `${template.thresholds[0]}/${template.thresholds[1]}`;
    const attackText = `+${template.attackModifier}`;
    const standardAttackText = `${template.standardAttack.name}: ${template.standardAttack.range}, ${template.standardAttack.damage}`;
    
    sheet.getRange(row, 1).setValue(adversaryName);
    sheet.getRange(row, 2).setValue(template.difficulty);
    sheet.getRange(row, 3).setValue(thresholdText);
    sheet.getRange(row, 4).setValue(attackText);
    sheet.getRange(row, 5).setValue(standardAttackText);
    
    for (let hpBox = 0; hpBox < template.hp; hpBox++) {
      const cell = sheet.getRange(row, 6 + hpBox);
      cell.insertCheckboxes();
      cell.setValue(false);
      cell.setHorizontalAlignment('center');
    }
    
    for (let stressBox = 0; stressBox < template.stress; stressBox++) {
      const cell = sheet.getRange(row, 6 + template.hp + stressBox);
      cell.insertCheckboxes();
      cell.setValue(false);
      cell.setHorizontalAlignment('center');
    }
    
    sheet.getRange(row, 6 + template.hp + template.stress).setValue('');
    row++;
  }
  
  return row;
}

function createAdversaryFeatures(sheet, row, totalCols, template) {
  if (!template.features || template.features.length === 0) {
    return row;
  }
  
  const featuresHeaderRange = sheet.getRange(row, 1, 1, totalCols);
  featuresHeaderRange.merge();
  featuresHeaderRange.setValue('Features & Abilities');
  featuresHeaderRange.setFontWeight('bold');
  featuresHeaderRange.setBackground('#e0e0e0');
  featuresHeaderRange.setFontColor('#424242');
  featuresHeaderRange.setHorizontalAlignment('center');
  row++;
  
  for (const feature of template.features) {
    const featureText = `${feature.name} - ${feature.type}: ${feature.description}`;
    const featureRange = sheet.getRange(row, 1, 1, totalCols);
    featureRange.merge();
    featureRange.setValue(featureText);
    featureRange.setWrap(true);
    featureRange.setBackground('#f5f5f5');
    featureRange.setFontColor('#424242');
    row++;
  }
  
  return row;
}

function createAdversaryExperience(sheet, row, totalCols, template) {
  if (!template.experience || template.experience.length === 0) {
    return row;
  }
  
  const expRange = sheet.getRange(row, 1, 1, totalCols);
  expRange.merge();
  expRange.setValue(`Experience: ${template.experience.join(', ')}`);
  expRange.setBackground('#f3e5f5');
  expRange.setFontColor('#4a148c');
  expRange.setFontStyle('italic');
  expRange.setFontWeight('bold');
  return row + 1;
}

function processEnvironmentBlock(sheet, row, selectedEnvironment) {
  const allEnvironments = getAllEnvironmentTemplates();
  const environment = allEnvironments[selectedEnvironment];
  
  if (!environment) return;
  
  createEnvironmentHeader(sheet, row, environment);
  row += 1;
  
  createEnvironmentType(sheet, row, environment);
  row += 1;
  
  if (environment.description) {
    createEnvironmentDescription(sheet, row, environment);
    row += 1;
  }
  
  if (environment.features && environment.features.length > 0) {
    createEnvironmentFeatures(sheet, row, environment);
  }
}

function createEnvironmentHeader(sheet, row, environment) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const envHeaderRange = sheet.getRange(row, 1, 1, maxCols);
  envHeaderRange.merge();
  envHeaderRange.setValue(`ENVIRONMENT: ${environment.name.toUpperCase()}`);
  envHeaderRange.setFontWeight('bold');
  envHeaderRange.setFontSize(16);
  envHeaderRange.setHorizontalAlignment('center');
  envHeaderRange.setBackground('#2e7d32');
  envHeaderRange.setFontColor('white');
}

function createEnvironmentType(sheet, row, environment) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const envTypeRange = sheet.getRange(row, 1, 1, maxCols);
  envTypeRange.merge();
  envTypeRange.setValue(`${environment.type} Environment`);
  envTypeRange.setFontWeight('bold');
  envTypeRange.setBackground('#c8e6c9');
  envTypeRange.setFontColor('#1b5e20');
  envTypeRange.setHorizontalAlignment('center');
}

function createEnvironmentDescription(sheet, row, environment) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const envDescRange = sheet.getRange(row, 1, 1, maxCols);
  envDescRange.merge();
  envDescRange.setValue(environment.description);
  envDescRange.setWrap(true);
  envDescRange.setBackground('#e8f5e8');
  envDescRange.setFontColor('#2e7d32');
  envDescRange.setFontStyle('italic');
}

function createEnvironmentFeatures(sheet, row, environment) {
  const maxCols = SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12;
  const envFeaturesHeaderRange = sheet.getRange(row, 1, 1, maxCols);
  envFeaturesHeaderRange.merge();
  envFeaturesHeaderRange.setValue('Environment Features');
  envFeaturesHeaderRange.setFontWeight('bold');
  envFeaturesHeaderRange.setBackground('#388e3c');
  envFeaturesHeaderRange.setFontColor('white');
  envFeaturesHeaderRange.setHorizontalAlignment('center');
  row++;
  
  for (const feature of environment.features) {
    const featureText = `${feature.name} - ${feature.type}: ${feature.description}`;
    const featureRange = sheet.getRange(row, 1, 1, maxCols);
    featureRange.merge();
    featureRange.setValue(featureText);
    featureRange.setWrap(true);
    featureRange.setBackground('#f1f8e9');
    featureRange.setFontColor('#2e7d32');
    row++;
  }
}

function formatEncounterSheet(sheet) {
  const formatting = SHEET_FORMATTING;
  
  sheet.setColumnWidth(1, formatting.NAME_COLUMN_WIDTH);
  sheet.setColumnWidth(2, formatting.DIFFICULTY_COLUMN_WIDTH);
  sheet.setColumnWidth(3, formatting.THRESHOLDS_COLUMN_WIDTH);
  sheet.setColumnWidth(4, formatting.ATTACK_COLUMN_WIDTH);
  sheet.setColumnWidth(5, formatting.STANDARD_ATTACK_COLUMN_WIDTH);
  
  const maxCol = sheet.getLastColumn();
  for (let col = 6; col <= maxCol - 1; col++) {
    sheet.setColumnWidth(col, formatting.CHECKBOX_COLUMN_WIDTH);
  }
  if (maxCol > 5) {
    sheet.setColumnWidth(maxCol, formatting.NOTES_COLUMN_WIDTH);
  }
}

function getAllEnvironmentTemplates() {
  // Check if cache is still valid
  if (DATA_CACHE.isDataValid() && DATA_CACHE.environments) {
    return DATA_CACHE.environments;
  }
  
  const coreEnvironments = loadEnvironmentsFromDrive('daggerheart-core-environments');
  const customEnvironments = loadEnvironmentsFromDrive('custom-environments');
  const allEnvironments = { ...coreEnvironments, ...customEnvironments };
  const filteredEnvironments = {};
  
  for (const [name, environment] of Object.entries(allEnvironments)) {
    if (!name.startsWith('//')) {
      filteredEnvironments[name] = environment;
    }
  }
  
  const result = Object.keys(filteredEnvironments).length === 0 
    ? getDefaultEnvironmentTemplates() 
    : filteredEnvironments;
  
  // Update cache
  DATA_CACHE.environments = result;
  DATA_CACHE.lastAccessTime = Date.now();
  
  return result;
}

function loadEnvironmentsFromDrive(fileName = 'daggerheart-core-environments') {
  try {
    const jsonFile = DATA_CACHE.getFileReference(fileName + '.json');
    if (!jsonFile) return {};
    
    const jsonContent = jsonFile.getBlob().getDataAsString();
    return JSON.parse(jsonContent);
    
  } catch (error) {
    console.error('Error loading environments from Drive:', error);
    return {};
  }
}

function getAllAdversaryTemplates() {
  // Check if cache is still valid
  if (DATA_CACHE.isDataValid() && DATA_CACHE.adversaries) {
    return DATA_CACHE.adversaries;
  }
  
  const coreAdversaries = loadAdversariesFromDrive('daggerheart-core-adversaries');
  const customAdversaries = loadAdversariesFromDrive('custom-adversaries');
  const allAdversaries = { ...coreAdversaries, ...customAdversaries };
  const filteredAdversaries = {};
  
  for (const [name, adversary] of Object.entries(allAdversaries)) {
    if (!name.startsWith('//')) {
      filteredAdversaries[name] = adversary;
    }
  }
  
  const result = Object.keys(filteredAdversaries).length === 0 
    ? getDefaultAdversaryTemplates() 
    : filteredAdversaries;
  
  // Update cache
  DATA_CACHE.adversaries = result;
  DATA_CACHE.lastAccessTime = Date.now();
  
  return result;
}

function loadAdversariesFromDrive(fileName = 'daggerheart-core-adversaries') {
  try {
    const jsonFile = DATA_CACHE.getFileReference(fileName + '.json');
    if (!jsonFile) return {};
    
    const jsonContent = jsonFile.getBlob().getDataAsString();
    return JSON.parse(jsonContent);
    
  } catch (error) {
    console.error('Error loading adversaries from Drive:', error);
    return {};
  }
}

function findDataFile(fileName) {
  return DATA_CACHE.getFileReference(fileName);
}

function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  
  sheet = spreadsheet.insertSheet(sheetName, 0);
  sheet.activate();
  
  return sheet;
}

function getBattleValue(adversaryType) {
  return BATTLE_VALUES[adversaryType] || 2;
}

function getOrCreateEncounterDataFolder() {
  // Use cached folder reference if available
  if (DATA_CACHE.folderReference) {
    return DATA_CACHE.folderReference;
  }
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(spreadsheet.getId());
  const parentFolders = file.getParents();
  
  let parentFolder;
  if (parentFolders.hasNext()) {
    parentFolder = parentFolders.next();
  } else {
    parentFolder = DriveApp.getRootFolder();
  }
  
  const existingEncounterDataFolders = parentFolder.getFoldersByName('Encounter Data');
  if (existingEncounterDataFolders.hasNext()) {
    DATA_CACHE.folderReference = existingEncounterDataFolders.next();
  } else {
    DATA_CACHE.folderReference = parentFolder.createFolder('Encounter Data');
  }
  
  return DATA_CACHE.folderReference;
}

function fetchAndParseAdversariesFromGitHub() {
  loadConfiguration(); // Only loads critical config
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const htmlOutput = createProgressDialog();
    ui.showModalDialog(htmlOutput, 'Importing Adversaries from GitHub');
    
    // Lazy load system timings
    const progressDelay = (ConfigManager.getSection('SYSTEM_TIMINGS')).PROGRESS_SPINNER_DELAY || 1000;
    Utilities.sleep(progressDelay);
    
    // Lazy load GitHub config
    const githubConfig = getGitHubConfig();
    const repoUrl = `${githubConfig.REPO_BASE}/${githubConfig.ADVERSARIES_PATH}`;
    const response = UrlFetchApp.fetch(repoUrl);
    const fileList = JSON.parse(response.getContentText());
    
    const adversaries = {};
    let processedCount = 0;
    
    // Lazy load rate limits
    const rateLimits = ConfigManager.getSection('GITHUB_RATE_LIMITS');
    const maxConcurrent = rateLimits.MAX_CONCURRENT_REQUESTS || 5;
    const requestDelay = rateLimits.REQUEST_DELAY_MS || 200;
    
    // Process files in batches to respect rate limits
    const markdownFiles = fileList.filter(file => file.name.endsWith('.md') && file.type === 'file');
    
    for (let i = 0; i < markdownFiles.length; i += maxConcurrent) {
      const batch = markdownFiles.slice(i, i + maxConcurrent);
      
      for (const file of batch) {
        try {
          const adversaryData = fetchAndParseAdversaryFileWithRetry(file.download_url, file.name);
          if (adversaryData) {
            adversaries[adversaryData.name] = adversaryData;
            processedCount++;
          }
        } catch (error) {
          console.error(`Error parsing ${file.name}:`, error);
        }
        
        if (requestDelay > 0) {
          Utilities.sleep(requestDelay);
        }
      }
      
      if (i + maxConcurrent < markdownFiles.length) {
        Utilities.sleep(requestDelay * 2);
      }
    }
    
    if (Object.keys(adversaries).length > 0) {
      saveAdversariesToDrive(adversaries, 'daggerheart-core-adversaries');
      const customFileExists = checkCustomAdversariesExists();
      ensureCustomAdversariesFile();
      
      DATA_CACHE.adversaries = null;
      DATA_CACHE.lastAccessTime = 0;
      
      let message = `Successfully imported ${processedCount} adversaries.\n\nSaved to Drive as 'daggerheart-core-adversaries.json'`;
      if (!customFileExists) {
        message += '\nAlso created \'custom-adversaries.json\' for your custom adversaries.';
      } else {
        message += '\nYour existing \'custom-adversaries.json\' was preserved.';
      }
      
      const completionDialog = createCompletionDialog(true, message);
      ui.showModalDialog(completionDialog, 'Import Complete');
    } else {
      const completionDialog = createCompletionDialog(false, 'No adversaries found or parsing failed.');
      ui.showModalDialog(completionDialog, 'Import Failed');
    }
    
  } catch (error) {
    console.error('Error fetching from GitHub:', error);
    const completionDialog = createCompletionDialog(false, 'Error importing adversaries: ' + error.message);
    ui.showModalDialog(completionDialog, 'Import Error');
  }
}

function fetchAndParseEnvironmentsFromGitHub() {
  loadConfiguration(); // Only loads critical config
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const htmlOutput = createProgressDialog('Importing Environments from GitHub', 'environment');
    ui.showModalDialog(htmlOutput, 'Importing Environments from GitHub');
    
    // Lazy load system timings
    const progressDelay = (ConfigManager.getSection('SYSTEM_TIMINGS')).PROGRESS_SPINNER_DELAY || 1000;
    Utilities.sleep(progressDelay);
    
    // Lazy load GitHub config
    const githubConfig = getGitHubConfig();
    const repoUrl = `${githubConfig.REPO_BASE}/${githubConfig.ENVIRONMENTS_PATH}`;
    const response = UrlFetchApp.fetch(repoUrl);
    const fileList = JSON.parse(response.getContentText());
    
    const environments = {};
    let processedCount = 0;
    
    // Lazy load rate limits
    const rateLimits = ConfigManager.getSection('GITHUB_RATE_LIMITS');
    const maxConcurrent = rateLimits.MAX_CONCURRENT_REQUESTS || 5;
    const requestDelay = rateLimits.REQUEST_DELAY_MS || 200;
    
    // Process files in batches to respect rate limits
    const markdownFiles = fileList.filter(file => file.name.endsWith('.md') && file.type === 'file');
    
    for (let i = 0; i < markdownFiles.length; i += maxConcurrent) {
      const batch = markdownFiles.slice(i, i + maxConcurrent);
      
      for (const file of batch) {
        try {
          const environmentData = fetchAndParseEnvironmentFileWithRetry(file.download_url, file.name);
          if (environmentData) {
            environments[environmentData.name] = environmentData;
            processedCount++;
          }
        } catch (error) {
          console.error(`Error parsing ${file.name}:`, error);
        }
        
        if (requestDelay > 0) {
          Utilities.sleep(requestDelay);
        }
      }
      
      if (i + maxConcurrent < markdownFiles.length) {
        Utilities.sleep(requestDelay * 2);
      }
    }
    
    if (Object.keys(environments).length > 0) {
      saveEnvironmentsToDrive(environments, 'daggerheart-core-environments');
      const customFileExists = checkCustomEnvironmentsExists();
      ensureCustomEnvironmentsFile();
      
      DATA_CACHE.environments = null;
      DATA_CACHE.lastAccessTime = 0;
      
      let message = `Successfully imported ${processedCount} environments.\n\nSaved to Drive as 'daggerheart-core-environments.json'`;
      if (!customFileExists) {
        message += '\nAlso created \'custom-environments.json\' for your custom environments.';
      } else {
        message += '\nYour existing \'custom-environments.json\' was preserved.';
      }
      
      const completionDialog = createCompletionDialog(true, message);
      ui.showModalDialog(completionDialog, 'Import Complete');
    } else {
      const completionDialog = createCompletionDialog(false, 'No environments found or parsing failed.');
      ui.showModalDialog(completionDialog, 'Import Failed');
    }
    
  } catch (error) {
    console.error('Error fetching environments from GitHub:', error);
    const completionDialog = createCompletionDialog(false, 'Error importing environments: ' + error.message);
    ui.showModalDialog(completionDialog, 'Import Error');
  }
}

function createProgressDialog(title = 'Importing Adversaries...', type = 'adversary') {
  // Lazy load dialog presets
  const dialogSize = getDialogPresets().PROGRESS_DIALOG || { WIDTH: 500, HEIGHT: 300 };
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 30px; 
            text-align: center;
            background: #f9f9f9;
          }
          .progress-container {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            max-width: 400px;
            margin: 0 auto;
          }
          .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #2196F3;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .progress-text {
            font-size: 18px;
            font-weight: bold;
            color: #333;
            margin-bottom: 10px;
          }
          .progress-detail {
            font-size: 14px;
            color: #666;
            line-height: 1.5;
          }
        </style>
      </head>
      <body>
        <div class="progress-container">
          <div class="spinner"></div>
          <div class="progress-text">${title}</div>
          <div class="progress-detail">
            Fetching ${type} data from whitwort/daggerheart-srd repository.<br>
            This may take a few moments to download and parse all ${type}s.
          </div>
          <div style="text-align: center; font-size: 10px; color: #999; margin-top: 20px; font-style: italic;">v0.2</div>
        </div>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html).setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
}


function createCompletionDialog(success, message) {
  // Lazy load dialog presets and timings
  const dialogSize = getDialogPresets().COMPLETION_DIALOG || { WIDTH: 600, HEIGHT: 400 };
  const autoCloseDelay = (ConfigManager.getSection('SYSTEM_TIMINGS')).AUTO_CLOSE_DELAY || 10000;
  const statusColor = success ? '#4CAF50' : '#f44336';
  const statusIcon = success ? '✅' : '❌';
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 30px; 
            text-align: center;
            background: #f9f9f9;
          }
          .completion-container {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            max-width: 500px;
            margin: 0 auto;
          }
          .status-icon {
            font-size: 48px;
            margin-bottom: 20px;
          }
          .status-text {
            font-size: 18px;
            font-weight: bold;
            color: ${statusColor};
            margin-bottom: 15px;
          }
          .message-text {
            font-size: 14px;
            color: #333;
            line-height: 1.5;
            margin-bottom: 25px;
          }
          .close-btn {
            background: ${statusColor};
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 6px;
            font-size: 16px;
            cursor: pointer;
          }
          .auto-close {
            font-size: 12px;
            color: #666;
            margin-top: 15px;
          }
        </style>
      </head>
      <body>
        <div class="completion-container">
          <div class="status-icon">${statusIcon}</div>
          <div class="status-text">${success ? 'Import Complete!' : 'Import Failed!'}</div>
          <div class="message-text">${message}</div>
          <button class="close-btn" onclick="google.script.host.close()">Close</button>
          <div class="auto-close">This dialog will close automatically in 10 seconds</div>
          <div style="text-align: center; font-size: 10px; color: #999; margin-top: 15px; font-style: italic;">v0.2</div>
        </div>
        
        <script>
          setTimeout(function() {
            google.script.host.close();
          }, ${autoCloseDelay});
        </script>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html).setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
}

function fetchAndParseAdversaryFileWithRetry(downloadUrl, fileName) {
  // Lazy load rate limits only when needed
  const maxRetries = (ConfigManager.getSection('GITHUB_RATE_LIMITS')).RETRY_ATTEMPTS || 3;
  const requestDelay = (ConfigManager.getSection('GITHUB_RATE_LIMITS')).REQUEST_DELAY_MS || 200;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(downloadUrl);
      const markdownContent = response.getContentText();
      return parseAdversaryMarkdown(markdownContent, fileName);
    } catch (error) {
      console.error(`Attempt ${attempt} failed for ${fileName}:`, error);
      
      if (attempt < maxRetries) {
        const retryDelay = requestDelay * attempt;
        Utilities.sleep(retryDelay);
      } else {
        throw error;
      }
    }
  }
  return null;
}

function fetchAndParseEnvironmentFileWithRetry(downloadUrl, fileName) {
  // Lazy load rate limits only when needed
  const maxRetries = (ConfigManager.getSection('GITHUB_RATE_LIMITS')).RETRY_ATTEMPTS || 3;
  const requestDelay = (ConfigManager.getSection('GITHUB_RATE_LIMITS')).REQUEST_DELAY_MS || 200;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(downloadUrl);
      const markdownContent = response.getContentText();
      return parseEnvironmentMarkdown(markdownContent, fileName);
    } catch (error) {
      console.error(`Attempt ${attempt} failed for ${fileName}:`, error);
      
      if (attempt < maxRetries) {
        const retryDelay = requestDelay * attempt;
        Utilities.sleep(retryDelay);
      } else {
        throw error;
      }
    }
  }
  return null;
}

function parseAdversaryMarkdown(markdown, fileName) {
  try {
    const nameMatch = markdown.match(/^#\s+(.+?)$/m);
    const name = nameMatch ? nameMatch[1].trim() : fileName.replace('.md', '').replace(/-/g, ' ');
    
    const adversary = {
      name: name,
      tier: 1,
      type: 'Standard',
      description: '',
      motives: '',
      difficulty: 12,
      thresholds: [5, 10],
      hp: 3,
      stress: 3,
      attackModifier: 0,
      standardAttack: {
        name: 'Basic Attack',
        range: 'Melee',
        damage: '1d6 phy'
      },
      experience: [],
      features: [],
      isCustom: false
    };
    
    const tierTypeMatch = markdown.match(/\*\*\*Tier\s+(\d+)\s+(.+?)\*\*\*/i);
    if (tierTypeMatch) {
      adversary.tier = parseInt(tierTypeMatch[1]);
      adversary.type = tierTypeMatch[2].trim();
    }
    
    const descMatch = markdown.match(/\*\*\*Tier[^*]+\*\*\*\s*\n\s*\*(.+?)\*\s*\n/i);
    if (descMatch) {
      adversary.description = descMatch[1].trim();
    }
    
    const motivesMatch = markdown.match(/\*\*Motives & Tactics:\*\*\s*(.+?)$/m);
    if (motivesMatch) {
      adversary.motives = motivesMatch[1].trim();
    }
    
    const statsBlockMatch = markdown.match(/>\s*\*\*Difficulty:\*\*\s*(\d+)\s*\|\s*\*\*Thresholds:\*\*\s*(\d+)\/(\d+)\s*\|\s*\*\*HP:\*\*\s*(\d+)\s*\|\s*\*\*Stress:\*\*\s*(\d+)/i);
    if (statsBlockMatch) {
      adversary.difficulty = parseInt(statsBlockMatch[1]);
      adversary.thresholds = [parseInt(statsBlockMatch[2]), parseInt(statsBlockMatch[3])];
      adversary.hp = parseInt(statsBlockMatch[4]);
      adversary.stress = parseInt(statsBlockMatch[5]);
    }
    
    const attackBlockMatch = markdown.match(/>\s*\*\*ATK:\*\*\s*([+-]?\d+)\s*\|\s*\*\*(.+?):\*\*\s*(.+?)\s*\|\s*(.+?)$/m);
    if (attackBlockMatch) {
      adversary.attackModifier = parseInt(attackBlockMatch[1]);
      adversary.standardAttack = {
        name: attackBlockMatch[2].trim(),
        range: attackBlockMatch[3].trim(),
        damage: attackBlockMatch[4].trim()
      };
    }
    
    const expMatch = markdown.match(/>\s*\*\*Experience:\*\*\s*(.+?)$/m);
    if (expMatch) {
      adversary.experience = expMatch[1].split(',').map(exp => exp.trim());
    }
    
    const featuresSection = markdown.match(/##\s*FEATURES\s*([\s\S]*?)(?=##|$)/i);
    if (featuresSection) {
      const featuresText = featuresSection[1];
      const featurePattern = /\*\*\*([^*]+?)\s*-\s*([^*]+?):\*\*\*([\s\S]*?)(?=\n\*\*\*|$)/g;
      const features = [];
      let match;
      
      while ((match = featurePattern.exec(featuresText)) !== null) {
        const name = match[1].trim();
        const type = match[2].trim();
        const description = match[3].trim();
        
        if (name && type && description) {
          features.push({
            name: name,
            type: type,
            description: description
          });
        }
      }
      
      adversary.features = features;
    }
    
    return adversary;
    
  } catch (error) {
    console.error(`Error parsing markdown for ${fileName}:`, error);
    return null;
  }
}

function parseEnvironmentMarkdown(markdown, fileName) {
  try {
    const nameMatch = markdown.match(/^#\s+(.+?)$/m);
    const name = nameMatch ? nameMatch[1].trim() : fileName.replace('.md', '').replace(/-/g, ' ');
    
    const environment = {
      name: name,
      type: 'Combat',
      description: '',
      features: []
    };
    
    const typeMatch = markdown.match(/\*\*\*.*?(Combat|Social|Traversal|Event).*?\*\*\*/i);
    if (typeMatch) {
      environment.type = typeMatch[1].trim();
    }
    
    const descMatch = markdown.match(/\*\*\*[^*]+\*\*\*\s*\n\s*\*(.+?)\*\s*\n/i);
    if (descMatch) {
      environment.description = descMatch[1].trim();
    }
    
    const featuresSection = markdown.match(/##\s*FEATURES\s*([\s\S]*?)(?=##|$)/i);
    if (featuresSection) {
      const featuresText = featuresSection[1];
      const featurePattern = /\*\*\*([^*]+?)\s*-\s*([^*]+?):\*\*\*([\s\S]*?)(?=\n\*\*\*|$)/g;
      const features = [];
      let match;
      
      while ((match = featurePattern.exec(featuresText)) !== null) {
        const name = match[1].trim();
        const type = match[2].trim();
        const description = match[3].trim();
        
        if (name && type && description) {
          features.push({
            name: name,
            type: type,
            description: description
          });
        }
      }
      
      environment.features = features;
    }
    
    return environment;
    
  } catch (error) {
    console.error(`Error parsing environment markdown for ${fileName}:`, error);
    return null;
  }
}

function saveAdversariesToDrive(adversaries, fileName) {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(adversaries, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName(fileName + '.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    } else {
      encounterDataFolder.createFile(fileName + '.json', jsonContent, MimeType.PLAIN_TEXT);
    }
    
  } catch (error) {
    console.error('Error saving to Drive:', error);
    throw new Error('Failed to save adversaries to Google Drive');
  }
}

function saveEnvironmentsToDrive(environments, fileName) {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(environments, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName(fileName + '.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    } else {
      encounterDataFolder.createFile(fileName + '.json', jsonContent, MimeType.PLAIN_TEXT);
    }
    
  } catch (error) {
    console.error('Error saving environments to Drive:', error);
    throw new Error('Failed to save environments to Google Drive');
  }
}

function ensureCustomAdversariesFile() {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const existingFiles = encounterDataFolder.getFilesByName('custom-adversaries.json');
    if (existingFiles.hasNext()) {
      return;
    }
    
    const emptyCustomFile = {
      "// Instructions": "Add your custom adversaries here. Copy the structure from core adversaries.",
      "// Example": {
        "name": "My Custom Enemy",
        "tier": 1,
        "type": "Standard",
        "description": "A custom adversary for my campaign",
        "motives": "Custom motives",
        "difficulty": 14,
        "thresholds": [8, 16],
        "hp": 4,
        "stress": 3,
        "attackModifier": 1,
        "standardAttack": {
          "name": "Custom Attack",
          "range": "Melee",
          "damage": "2d6+2 phy"
        },
        "experience": ["Custom Skill +2"],
        "features": [
          {
            "name": "Custom Feature",
            "type": "Action",
            "description": "Description of the custom feature"
          }
        ],
        "isCustom": true
      }
    };
    
    const jsonContent = JSON.stringify(emptyCustomFile, null, 2);
    encounterDataFolder.createFile('custom-adversaries.json', jsonContent, MimeType.PLAIN_TEXT);
    
  } catch (error) {
    console.error('Error creating custom adversaries file:', error);
  }
}

function ensureCustomEnvironmentsFile() {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const existingFiles = encounterDataFolder.getFilesByName('custom-environments.json');
    if (existingFiles.hasNext()) {
      return;
    }
    
    const emptyCustomFile = {
      "// Instructions": "Add your custom environments here. Copy the structure from core environments.",
      "// Example": {
        "name": "My Custom Environment",
        "type": "Combat",
        "description": "A custom environment for my campaign",
        "features": [
          {
            "name": "Custom Feature",
            "type": "Action",
            "description": "Description of the custom environment feature"
          }
        ]
      }
    };
    
    const jsonContent = JSON.stringify(emptyCustomFile, null, 2);
    encounterDataFolder.createFile('custom-environments.json', jsonContent, MimeType.PLAIN_TEXT);
    
  } catch (error) {
    console.error('Error creating custom environments file:', error);
  }
}

function checkCustomAdversariesExists() {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const files = encounterDataFolder.getFilesByName('custom-adversaries.json');
    return files.hasNext();
  } catch (error) {
    console.error('Error checking custom adversaries file:', error);
    return false;
  }
}

function checkCustomEnvironmentsExists() {
  try {
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const files = encounterDataFolder.getFilesByName('custom-environments.json');
    return files.hasNext();
  } catch (error) {
    console.error('Error checking custom environments file:', error);
    return false;
  }
}

function getDefaultEnvironmentTemplates() {
  return {
    "Example Combat Environment": {
      name: "Example Combat Environment",
      type: "Combat",
      description: "A basic combat environment for encounters",
      features: [
        {
          name: "Dangerous Terrain",
          type: "Passive",
          description: "Certain areas of the battlefield are hazardous"
        }
      ]
    }
  };
}

function getDefaultAdversaryTemplates() {
  return {
    "Bandit": {
      name: "Bandit",
      tier: 1,
      type: "Standard",
      description: "A common highway robber",
      motives: "Rob travelers, avoid guards",
      difficulty: 12,
      thresholds: [5, 10],
      hp: 3,
      stress: 3,
      attackModifier: 1,
      standardAttack: {
        name: "Shortsword",
        range: "Melee",
        damage: "1d8+1 phy"
      },
      experience: ["Stealth +2"],
      features: [],
      isCustom: false
    }
  };
}

// Custom Editor Functions
function showCustomAdversaryEditor() {
  loadConfiguration(); // Only loads critical config
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = createCustomAdversaryEditor();
  
  // Lazy load dialog presets only when needed
  const dialogSize = getDialogPresets().ADVERSARY_EDITOR || { WIDTH: 1000, HEIGHT: 800 };
  htmlOutput.setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
  ui.showModalDialog(htmlOutput, 'Custom Adversary Editor - v0.2');
}

function showCustomEnvironmentEditor() {
  loadConfiguration(); // Only loads critical config
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = createCustomEnvironmentEditor();
  
  // Lazy load dialog presets only when needed
  const dialogSize = getDialogPresets().ENVIRONMENT_EDITOR || { WIDTH: 1000, HEIGHT: 800 };
  htmlOutput.setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
  ui.showModalDialog(htmlOutput, 'Custom Environment Editor - v0.2');
}

function createCustomAdversaryEditor() {
  const customAdversaries = loadAdversariesFromDrive('custom-adversaries');
  const coreAdversaries = loadAdversariesFromDrive('daggerheart-core-adversaries');
  
  const customNames = Object.keys(customAdversaries)
    .filter(name => !name.startsWith('//'))
    .sort();
  const coreNames = Object.keys(coreAdversaries)
    .filter(name => !name.startsWith('//'))
    .sort();
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          ${getEditorCSS()}
        </style>
      </head>
      <body>
        <div class="editor-container">
          ${generateAdversaryEditorHTML(customNames, coreNames)}
        </div>
        <script>
          ${generateAdversaryEditorScript(customAdversaries, coreAdversaries)}
        </script>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

function createCustomEnvironmentEditor() {
  const customEnvironments = loadEnvironmentsFromDrive('custom-environments');
  const coreEnvironments = loadEnvironmentsFromDrive('daggerheart-core-environments');
  
  const customNames = Object.keys(customEnvironments)
    .filter(name => !name.startsWith('//'))
    .sort();
  const coreNames = Object.keys(coreEnvironments)
    .filter(name => !name.startsWith('//'))
    .sort();
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          ${getEditorCSS()}
        </style>
      </head>
      <body>
        <div class="editor-container">
          ${generateEnvironmentEditorHTML(customNames, coreNames)}
        </div>
        <script>
          ${generateEnvironmentEditorScript(customEnvironments, coreEnvironments)}
        </script>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

function getEditorCSS() {
  return `
    * { box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 20px; 
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      min-height: 100vh;
    }
    .editor-container {
      background: white;
      padding: 30px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      max-width: 1200px;
      margin: 0 auto;
      height: calc(100vh - 40px);
      overflow-y: auto;
    }
    .editor-header {
      text-align: center;
      margin-bottom: 30px;
      padding-bottom: 20px;
      border-bottom: 3px solid #e3f2fd;
    }
    .editor-header h2 {
      margin: 0;
      color: #1976d2;
      font-size: 24px;
      font-weight: 700;
    }
    .editor-controls {
      display: flex;
      gap: 15px;
      margin-bottom: 30px;
      align-items: center;
      flex-wrap: wrap;
    }
    .control-group {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .control-group label {
      font-weight: 600;
      color: #333;
    }
    select, input, textarea {
      padding: 10px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: all 0.3s ease;
    }
    select:focus, input:focus, textarea:focus {
      border-color: #1976d2;
      outline: none;
      box-shadow: 0 0 0 3px rgba(25,118,210,0.2);
    }
    .btn {
      padding: 10px 20px;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    .btn-primary {
      background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
      color: white;
    }
    .btn-success {
      background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
      color: white;
    }
    .btn-success:disabled {
      background: #bdbdbd;
      cursor: not-allowed;
    }
    .btn-danger {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
      color: white;
    }
    .btn-secondary {
      background: linear-gradient(135deg, #757575 0%, #616161 100%);
      color: white;
    }
    .form-section {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #e0e0e0;
      margin-bottom: 20px;
    }
    .form-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      margin-bottom: 20px;
    }
    .form-section h3 {
      margin: 0 0 15px 0;
      color: #1976d2;
      font-size: 18px;
      font-weight: 600;
    }
    .form-row {
      margin-bottom: 15px;
    }
    .form-row:last-child {
      margin-bottom: 0;
    }
    .form-row label {
      display: block;
      font-weight: 600;
      color: #333;
      margin-bottom: 5px;
    }
    .form-row input, .form-row select, .form-row textarea {
      width: 100%;
    }
    .form-row textarea {
      min-height: 80px;
      resize: vertical;
    }
    .thresholds-input {
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .thresholds-input input {
      flex: 1;
    }
    .attack-section {
      background: #e8f5e8;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #4caf50;
      margin-bottom: 20px;
    }
    .attack-section h3 {
      margin: 0 0 15px 0;
      color: #2e7d32;
      font-size: 18px;
      font-weight: 600;
    }
    .attack-grid {
      display: grid;
      grid-template-columns: 1fr 2fr 1fr;
      gap: 15px;
    }
    .features-section {
      background: #f5f5f5;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #757575;
      margin-bottom: 20px;
    }
    .features-section h3 {
      margin: 0 0 15px 0;
      color: #424242;
      font-size: 18px;
      font-weight: 600;
    }
    .feature-item {
      background: white;
      padding: 15px;
      border-radius: 8px;
      border: 1px solid #ddd;
      margin-bottom: 10px;
      position: relative;
    }
    .feature-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }
    .feature-controls {
      display: flex;
      gap: 10px;
    }
    .feature-grid {
      display: grid;
      grid-template-columns: 2fr 1fr;
      gap: 15px;
      margin-bottom: 10px;
    }
    .experience-section {
      background: #f3e5f5;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #9c27b0;
      margin-bottom: 20px;
    }
    .experience-section h3 {
      margin: 0 0 15px 0;
      color: #4a148c;
      font-size: 18px;
      font-weight: 600;
    }
    .experience-list {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 15px;
    }
    .experience-tag {
      background: #9c27b0;
      color: white;
      padding: 5px 10px;
      border-radius: 20px;
      font-size: 12px;
      display: flex;
      align-items: center;
      gap: 5px;
    }
    .experience-tag .remove {
      background: rgba(255,255,255,0.3);
      border-radius: 50%;
      width: 16px;
      height: 16px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 10px;
    }
    .experience-input {
      display: flex;
      gap: 10px;
    }
    .experience-input input {
      flex: 1;
    }
    .action-buttons {
      display: flex;
      gap: 15px;
      justify-content: center;
      margin-top: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .version-info {
      text-align: center;
      font-size: 11px;
      color: #999;
      margin-top: 20px;
      font-style: italic;
    }
    .hidden {
      display: none;
    }
    .name-warning {
      margin-top: 5px;
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
    }
    .name-warning.overwrite {
      background: #ffcdd2;
      color: #b71c1c;
      border: 1px solid #f44336;
    }
    .name-warning.override {
      background: #fff3c4;
      color: #e65100;
      border: 1px solid #ff9800;
    }
  `;
}

function generateAdversaryEditorHTML(customNames, coreNames) {
  const validation = FORM_VALIDATION;
  
  let selectOptions = '<option value="">Create New Adversary</option>';
  
  if (customNames.length > 0) {
    selectOptions += '<optgroup label="Custom Adversaries">';
    customNames.forEach(name => {
      selectOptions += `<option value="${name}">${name}</option>`;
    });
    selectOptions += '</optgroup>';
  }
  
  if (coreNames.length > 0) {
    selectOptions += '<optgroup label="Core Adversaries">';
    coreNames.forEach(name => {
      selectOptions += `<option value="${name}">${name}</option>`;
    });
    selectOptions += '</optgroup>';
  }

  return `
    <div class="editor-header">
      <h2>Custom Adversary Editor</h2>
    </div>
    
    <div class="editor-controls">
      <div class="control-group">
        <label>Adversary:</label>
        <select id="adversary-select" onchange="loadAdversary()">
          ${selectOptions}
        </select>
      </div>
      <button class="btn btn-secondary" onclick="clearForm()">Clear Form</button>
      <button class="btn btn-danger" onclick="deleteAdversary()" id="delete-btn" style="display: none;">Delete Adversary</button>
    </div>
    
    <div class="form-grid">
      <div class="form-section">
        <h3>Basic Information</h3>
        <div class="form-row">
          <label>Name:</label>
          <input type="text" id="adversary-name" placeholder="Enter adversary name" 
                 maxlength="${validation.MAX_NAME_LENGTH}" oninput="checkNameWarnings()">
          <div id="name-warning" class="name-warning hidden"></div>
        </div>
        <div class="form-row">
          <label>Tier:</label>
          <select id="adversary-tier">
            <option value="1">Tier 1</option>
            <option value="2">Tier 2</option>
            <option value="3">Tier 3</option>
            <option value="4">Tier 4</option>
          </select>
        </div>
        <div class="form-row">
          <label>Type:</label>
          <select id="adversary-type">
            <option value="Minion">Minion</option>
            <option value="Standard">Standard</option>
            <option value="Horde">Horde</option>
            <option value="Skulk">Skulk</option>
            <option value="Ranged">Ranged</option>
            <option value="Support">Support</option>
            <option value="Social">Social</option>
            <option value="Leader">Leader</option>
            <option value="Bruiser">Bruiser</option>
            <option value="Solo">Solo</option>
          </select>
        </div>
        <div class="form-row">
          <label>Description:</label>
          <textarea id="adversary-description" placeholder="Describe the adversary's appearance and demeanor"></textarea>
        </div>
        <div class="form-row">
          <label>Motives & Tactics:</label>
          <textarea id="adversary-motives" placeholder="Describe the adversary's goals and behavior"></textarea>
        </div>
      </div>
      
      <div class="form-section">
        <h3>Combat Statistics</h3>
        <div class="form-row">
          <label>Difficulty:</label>
          <input type="number" id="adversary-difficulty" 
                 min="${validation.MIN_DIFFICULTY}" max="${validation.MAX_DIFFICULTY}" value="12">
        </div>
        <div class="form-row">
          <label>Thresholds:</label>
          <div class="thresholds-input">
            <input type="number" id="threshold-major" 
                   min="${validation.MIN_THRESHOLD_VALUE}" max="${validation.MAX_THRESHOLD_VALUE}" 
                   value="5" placeholder="Major">
            <span>/</span>
            <input type="number" id="threshold-severe" 
                   min="${validation.MIN_THRESHOLD_VALUE}" max="${validation.MAX_THRESHOLD_VALUE}" 
                   value="10" placeholder="Severe">
          </div>
        </div>
        <div class="form-row">
          <label>Hit Points:</label>
          <input type="number" id="adversary-hp" 
                 min="${validation.MIN_HP_VALUE}" max="${validation.MAX_HP_VALUE}" value="3">
        </div>
        <div class="form-row">
          <label>Stress:</label>
          <input type="number" id="adversary-stress" 
                 min="${validation.MIN_STRESS_VALUE}" max="${validation.MAX_STRESS_VALUE}" value="3">
        </div>
        <div class="form-row">
          <label>Attack Modifier:</label>
          <input type="number" id="attack-modifier" 
                 min="${validation.MIN_ATTACK_MODIFIER}" max="${validation.MAX_ATTACK_MODIFIER}" value="0">
        </div>
      </div>
    </div>
    
    <div class="attack-section">
      <h3>Standard Attack</h3>
      <div class="attack-grid">
        <div class="form-row">
          <label>Attack Name:</label>
          <input type="text" id="attack-name" placeholder="e.g., Shortsword" value="Basic Attack">
        </div>
        <div class="form-row">
          <label>Range:</label>
          <select id="attack-range">
            <option value="Melee">Melee</option>
            <option value="Very Close">Very Close</option>
            <option value="Close">Close</option>
            <option value="Far">Far</option>
            <option value="Very Far">Very Far</option>
          </select>
        </div>
        <div class="form-row">
          <label>Damage:</label>
          <input type="text" id="attack-damage" placeholder="e.g., 1d8+1 phy" value="1d6 phy">
        </div>
      </div>
    </div>
    
    <div class="experience-section">
      <h3>Experience</h3>
      <div class="experience-list" id="experience-list"></div>
      <div class="experience-input">
        <input type="text" id="new-experience" placeholder="e.g., Stealth +2">
        <button class="btn btn-primary" onclick="addExperience()">Add Experience</button>
      </div>
    </div>
    
    <div class="features-section">
      <h3>Features & Abilities</h3>
      <div id="features-list"></div>
      <button class="btn btn-primary" onclick="addFeature()">Add Feature</button>
    </div>
    
    <div class="action-buttons">
      <button class="btn btn-success" onclick="saveAdversary()" id="save-btn">Save Adversary</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </div>
    
    <div class="version-info">Custom Adversary Editor v.0.2</div>
  `;
}

function generateEnvironmentEditorHTML(customNames, coreNames) {
  let selectOptions = '<option value="">Create New Environment</option>';
  
  if (customNames.length > 0) {
    selectOptions += '<optgroup label="Custom Environments">';
    customNames.forEach(name => {
      selectOptions += `<option value="${name}">${name}</option>`;
    });
    selectOptions += '</optgroup>';
  }
  
  if (coreNames.length > 0) {
    selectOptions += '<optgroup label="Core Environments">';
    coreNames.forEach(name => {
      selectOptions += `<option value="${name}">${name}</option>`;
    });
    selectOptions += '</optgroup>';
  }

  return `
    <div class="editor-header">
      <h2>Custom Environment Editor</h2>
    </div>
    
    <div class="editor-controls">
      <div class="control-group">
        <label>Environment:</label>
        <select id="environment-select" onchange="loadEnvironment()">
          ${selectOptions}
        </select>
      </div>
      <button class="btn btn-secondary" onclick="clearForm()">Clear Form</button>
      <button class="btn btn-danger" onclick="deleteEnvironment()" id="delete-btn" style="display: none;">Delete Environment</button>
    </div>
    
    <div class="form-section">
      <h3>Basic Information</h3>
      <div class="form-row">
        <label>Name:</label>
        <input type="text" id="environment-name" placeholder="Enter environment name" oninput="checkNameWarnings()">
        <div id="name-warning" class="name-warning hidden"></div>
      </div>
      <div class="form-row">
        <label>Type:</label>
        <select id="environment-type">
          <option value="Combat">Combat</option>
          <option value="Social">Social</option>
          <option value="Traversal">Traversal</option>
          <option value="Event">Event</option>
        </select>
      </div>
      <div class="form-row">
        <label>Description:</label>
        <textarea id="environment-description" placeholder="Describe the environment and its atmosphere"></textarea>
      </div>
    </div>
    
    <div class="features-section" style="background: #e8f5e8; border-color: #4caf50;">
      <h3 style="color: #2e7d32;">Environment Features</h3>
      <div id="features-list"></div>
      <button class="btn btn-primary" onclick="addFeature()">Add Feature</button>
    </div>
    
    <div class="action-buttons">
      <button class="btn btn-success" onclick="saveEnvironment()" id="save-btn">Save Environment</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </div>
    
    <div class="version-info">Custom Environment Editor v.0.2</div>
  `;
}

function generateAdversaryEditorScript(customAdversaries, coreAdversaries) {
  return `
    const customAdversariesData = ${JSON.stringify(customAdversaries)};
    const coreAdversariesData = ${JSON.stringify(coreAdversaries)};
    const allAdversariesData = { ...coreAdversariesData, ...customAdversariesData };
    let currentAdversary = null;
    let experienceList = [];
    let featuresList = [];
    
    function checkNameWarnings() {
      const nameInput = document.getElementById('adversary-name');
      const warningDiv = document.getElementById('name-warning');
      const name = nameInput.value.trim();
      
      if (!name) {
        warningDiv.className = 'name-warning hidden';
        return;
      }
      
      if (customAdversariesData[name]) {
        warningDiv.textContent = '⚠️ This will overwrite the existing custom adversary with this name.';
        warningDiv.className = 'name-warning overwrite';
      } else if (coreAdversariesData[name]) {
        warningDiv.textContent = '⚠️ This custom adversary will override the core adversary with this name.';
        warningDiv.className = 'name-warning override';
      } else {
        warningDiv.className = 'name-warning hidden';
      }
    }
    
    function loadAdversary() {
      const selectElement = document.getElementById('adversary-select');
      const adversaryName = selectElement.value;
      const deleteBtn = document.getElementById('delete-btn');
      
      if (adversaryName && allAdversariesData[adversaryName]) {
        currentAdversary = allAdversariesData[adversaryName];
        populateForm(currentAdversary);
        deleteBtn.style.display = customAdversariesData[adversaryName] ? 'inline-block' : 'none';
      } else {
        clearForm();
        deleteBtn.style.display = 'none';
      }
      
      checkNameWarnings();
    }
    
    function populateForm(adversary) {
      document.getElementById('adversary-name').value = adversary.name || '';
      document.getElementById('adversary-tier').value = adversary.tier || 1;
      document.getElementById('adversary-type').value = adversary.type || 'Standard';
      document.getElementById('adversary-description').value = adversary.description || '';
      document.getElementById('adversary-motives').value = adversary.motives || '';
      document.getElementById('adversary-difficulty').value = adversary.difficulty || 12;
      document.getElementById('threshold-major').value = adversary.thresholds ? adversary.thresholds[0] : 5;
      document.getElementById('threshold-severe').value = adversary.thresholds ? adversary.thresholds[1] : 10;
      document.getElementById('adversary-hp').value = adversary.hp || 3;
      document.getElementById('adversary-stress').value = adversary.stress || 3;
      document.getElementById('attack-modifier').value = adversary.attackModifier || 0;
      
      if (adversary.standardAttack) {
        document.getElementById('attack-name').value = adversary.standardAttack.name || 'Basic Attack';
        document.getElementById('attack-range').value = adversary.standardAttack.range || 'Melee';
        document.getElementById('attack-damage').value = adversary.standardAttack.damage || '1d6 phy';
      }
      
      experienceList = adversary.experience ? [...adversary.experience] : [];
      featuresList = adversary.features ? adversary.features.map(f => ({...f})) : [];
      
      updateExperienceDisplay();
      updateFeaturesDisplay();
    }
    
    function clearForm() {
      document.getElementById('adversary-select').value = '';
      document.getElementById('adversary-name').value = '';
      document.getElementById('adversary-tier').value = '1';
      document.getElementById('adversary-type').value = 'Standard';
      document.getElementById('adversary-description').value = '';
      document.getElementById('adversary-motives').value = '';
      document.getElementById('adversary-difficulty').value = '12';
      document.getElementById('threshold-major').value = '5';
      document.getElementById('threshold-severe').value = '10';
      document.getElementById('adversary-hp').value = '3';
      document.getElementById('adversary-stress').value = '3';
      document.getElementById('attack-modifier').value = '0';
      document.getElementById('attack-name').value = 'Basic Attack';
      document.getElementById('attack-range').value = 'Melee';
      document.getElementById('attack-damage').value = '1d6 phy';
      
      experienceList = [];
      featuresList = [];
      currentAdversary = null;
      
      updateExperienceDisplay();
      updateFeaturesDisplay();
      
      document.getElementById('delete-btn').style.display = 'none';
      document.getElementById('name-warning').className = 'name-warning hidden';
    }
    
    function addExperience() {
      const input = document.getElementById('new-experience');
      const experience = input.value.trim();
      
      if (experience && !experienceList.includes(experience)) {
        experienceList.push(experience);
        input.value = '';
        updateExperienceDisplay();
      }
    }
    
    function removeExperience(index) {
      experienceList.splice(index, 1);
      updateExperienceDisplay();
    }
    
    function updateExperienceDisplay() {
      const container = document.getElementById('experience-list');
      container.innerHTML = experienceList.map((exp, index) => 
        \`<div class="experience-tag">
          \${exp}
          <div class="remove" onclick="removeExperience(\${index})">×</div>
        </div>\`
      ).join('');
    }
    
    function addFeature() {
      featuresList.push({
        name: '',
        type: 'Passive',
        description: ''
      });
      updateFeaturesDisplay();
    }
    
    function removeFeature(index) {
      featuresList.splice(index, 1);
      updateFeaturesDisplay();
    }
    
    function updateFeature(index, field, value) {
      if (featuresList[index]) {
        featuresList[index][field] = value;
      }
    }
    
    function updateFeaturesDisplay() {
      const container = document.getElementById('features-list');
      container.innerHTML = featuresList.map((feature, index) => 
        \`<div class="feature-item">
          <div class="feature-header">
            <strong>Feature \${index + 1}</strong>
            <div class="feature-controls">
              <button class="btn btn-danger" onclick="removeFeature(\${index})">Remove</button>
            </div>
          </div>
          <div class="feature-grid">
            <div class="form-row">
              <label>Feature Name:</label>
              <input type="text" value="\${feature.name}" placeholder="e.g., Relentless" 
                     onchange="updateFeature(\${index}, 'name', this.value)">
            </div>
            <div class="form-row">
              <label>Type:</label>
              <select onchange="updateFeature(\${index}, 'type', this.value)">
                <option value="Passive" \${feature.type === 'Passive' ? 'selected' : ''}>Passive</option>
                <option value="Action" \${feature.type === 'Action' ? 'selected' : ''}>Action</option>
                <option value="Reaction" \${feature.type === 'Reaction' ? 'selected' : ''}>Reaction</option>
                <option value="Action: Countdown" \${feature.type === 'Action: Countdown' ? 'selected' : ''}>Action: Countdown</option>
                <option value="Other" \${feature.type === 'Other' ? 'selected' : ''}>Other</option>
              </select>
            </div>
          </div>
          <div class="form-row">
            <label>Description:</label>
            <textarea value="\${feature.description}" placeholder="Describe what this feature does" 
                      onchange="updateFeature(\${index}, 'description', this.value)">\${feature.description}</textarea>
          </div>
        </div>\`
      ).join('');
    }
    
    function saveAdversary() {
      const name = document.getElementById('adversary-name').value.trim();
      
      if (!name) {
        alert('Please enter an adversary name.');
        return;
      }
      
      const saveBtn = document.getElementById('save-btn');
      saveBtn.innerHTML = 'Saving...';
      saveBtn.disabled = true;
      
      const adversary = {
        name: name,
        tier: parseInt(document.getElementById('adversary-tier').value),
        type: document.getElementById('adversary-type').value,
        description: document.getElementById('adversary-description').value.trim(),
        motives: document.getElementById('adversary-motives').value.trim(),
        difficulty: parseInt(document.getElementById('adversary-difficulty').value),
        thresholds: [
          parseInt(document.getElementById('threshold-major').value),
          parseInt(document.getElementById('threshold-severe').value)
        ],
        hp: parseInt(document.getElementById('adversary-hp').value),
        stress: parseInt(document.getElementById('adversary-stress').value),
        attackModifier: parseInt(document.getElementById('attack-modifier').value),
        standardAttack: {
          name: document.getElementById('attack-name').value.trim(),
          range: document.getElementById('attack-range').value,
          damage: document.getElementById('attack-damage').value.trim()
        },
        experience: [...experienceList],
        features: featuresList.filter(f => f.name.trim() && f.description.trim()),
        isCustom: true
      };
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Adversary saved successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error saving adversary: ' + error.message);
          saveBtn.innerHTML = 'Save Adversary';
          saveBtn.disabled = false;
        })
        .saveCustomAdversary(adversary);
    }
    
    function deleteAdversary() {
      const name = document.getElementById('adversary-name').value.trim();
      
      if (!name) {
        alert('No adversary selected for deletion.');
        return;
      }
      
      if (!customAdversariesData[name]) {
        alert('You can only delete custom adversaries, not core adversaries.');
        return;
      }
      
      if (confirm(\`Are you sure you want to delete the custom adversary "\${name}"? This action cannot be undone.\`)) {
        google.script.run
          .withSuccessHandler(() => {
            alert('Adversary deleted successfully!');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error deleting adversary: ' + error.message);
          })
          .deleteCustomAdversary(name);
      }
    }
    
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' && e.target.id === 'new-experience') {
        e.preventDefault();
        addExperience();
      }
    });
    
    updateExperienceDisplay();
    updateFeaturesDisplay();
  `;
}

function generateEnvironmentEditorScript(customEnvironments, coreEnvironments) {
  return `
    const customEnvironmentsData = ${JSON.stringify(customEnvironments)};
    const coreEnvironmentsData = ${JSON.stringify(coreEnvironments)};
    const allEnvironmentsData = { ...coreEnvironmentsData, ...customEnvironmentsData };
    let currentEnvironment = null;
    let featuresList = [];
    
    function checkNameWarnings() {
      const nameInput = document.getElementById('environment-name');
      const warningDiv = document.getElementById('name-warning');
      const name = nameInput.value.trim();
      
      if (!name) {
        warningDiv.className = 'name-warning hidden';
        return;
      }
      
      if (customEnvironmentsData[name]) {
        warningDiv.textContent = '⚠️ This will overwrite the existing custom environment with this name.';
        warningDiv.className = 'name-warning overwrite';
      } else if (coreEnvironmentsData[name]) {
        warningDiv.textContent = '⚠️ This custom environment will override the core environment with this name.';
        warningDiv.className = 'name-warning override';
      } else {
        warningDiv.className = 'name-warning hidden';
      }
    }
    
    function loadEnvironment() {
      const selectElement = document.getElementById('environment-select');
      const environmentName = selectElement.value;
      const deleteBtn = document.getElementById('delete-btn');
      
      if (environmentName && allEnvironmentsData[environmentName]) {
        currentEnvironment = allEnvironmentsData[environmentName];
        populateForm(currentEnvironment);
        deleteBtn.style.display = customEnvironmentsData[environmentName] ? 'inline-block' : 'none';
      } else {
        clearForm();
        deleteBtn.style.display = 'none';
      }
      
      checkNameWarnings();
    }
    
    function populateForm(environment) {
      document.getElementById('environment-name').value = environment.name || '';
      document.getElementById('environment-type').value = environment.type || 'Combat';
      document.getElementById('environment-description').value = environment.description || '';
      
      featuresList = environment.features ? environment.features.map(f => ({...f})) : [];
      
      updateFeaturesDisplay();
    }
    
    function clearForm() {
      document.getElementById('environment-select').value = '';
      document.getElementById('environment-name').value = '';
      document.getElementById('environment-type').value = 'Combat';
      document.getElementById('environment-description').value = '';
      
      featuresList = [];
      currentEnvironment = null;
      
      updateFeaturesDisplay();
      
      document.getElementById('delete-btn').style.display = 'none';
      document.getElementById('name-warning').className = 'name-warning hidden';
    }
    
    function addFeature() {
      featuresList.push({
        name: '',
        type: 'Passive',
        description: ''
      });
      updateFeaturesDisplay();
    }
    
    function removeFeature(index) {
      featuresList.splice(index, 1);
      updateFeaturesDisplay();
    }
    
    function updateFeature(index, field, value) {
      if (featuresList[index]) {
        featuresList[index][field] = value;
      }
    }
    
    function updateFeaturesDisplay() {
      const container = document.getElementById('features-list');
      container.innerHTML = featuresList.map((feature, index) => 
        \`<div class="feature-item">
          <div class="feature-header">
            <strong>Feature \${index + 1}</strong>
            <div class="feature-controls">
              <button class="btn btn-danger" onclick="removeFeature(\${index})">Remove</button>
            </div>
          </div>
          <div class="feature-grid">
            <div class="form-row">
              <label>Feature Name:</label>
              <input type="text" value="\${feature.name}" placeholder="e.g., Dangerous Terrain" 
                     onchange="updateFeature(\${index}, 'name', this.value)">
            </div>
            <div class="form-row">
              <label>Type:</label>
              <select onchange="updateFeature(\${index}, 'type', this.value)">
                <option value="Passive" \${feature.type === 'Passive' ? 'selected' : ''}>Passive</option>
                <option value="Action" \${feature.type === 'Action' ? 'selected' : ''}>Action</option>
                <option value="Reaction" \${feature.type === 'Reaction' ? 'selected' : ''}>Reaction</option>
              </select>
            </div>
          </div>
          <div class="form-row">
            <label>Description:</label>
            <textarea value="\${feature.description}" placeholder="Describe what this environment feature does" 
                      onchange="updateFeature(\${index}, 'description', this.value)">\${feature.description}</textarea>
          </div>
        </div>\`
      ).join('');
    }
    
    function saveEnvironment() {
      const name = document.getElementById('environment-name').value.trim();
      
      if (!name) {
        alert('Please enter an environment name.');
        return;
      }
      
      const saveBtn = document.getElementById('save-btn');
      saveBtn.innerHTML = 'Saving...';
      saveBtn.disabled = true;
      
      const environment = {
        name: name,
        type: document.getElementById('environment-type').value,
        description: document.getElementById('environment-description').value.trim(),
        features: featuresList.filter(f => f.name.trim() && f.description.trim())
      };
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Environment saved successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error saving environment: ' + error.message);
          saveBtn.innerHTML = 'Save Environment';
          saveBtn.disabled = false;
        })
        .saveCustomEnvironment(environment);
    }
    
    function deleteEnvironment() {
      const name = document.getElementById('environment-name').value.trim();
      
      if (!name) {
        alert('No environment selected for deletion.');
        return;
      }
      
      if (!customEnvironmentsData[name]) {
        alert('You can only delete custom environments, not core environments.');
        return;
      }
      
      if (confirm(\`Are you sure you want to delete the custom environment "\${name}"? This action cannot be undone.\`)) {
        google.script.run
          .withSuccessHandler(() => {
            alert('Environment deleted successfully!');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error deleting environment: ' + error.message);
          })
          .deleteCustomEnvironment(name);
      }
    }
    
    updateFeaturesDisplay();
  `;
}

// Additional utility functions for custom content management
function saveCustomAdversary(adversary) {
  try {
    const customAdversaries = loadAdversariesFromDrive('custom-adversaries');
    customAdversaries[adversary.name] = adversary;
    
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(customAdversaries, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName('custom-adversaries.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    } else {
      encounterDataFolder.createFile('custom-adversaries.json', jsonContent, MimeType.PLAIN_TEXT);
    }
    
    DATA_CACHE.adversaries = null;
    DATA_CACHE.lastAccessTime = 0;
    
    return { success: true };
  } catch (error) {
    console.error('Error saving custom adversary:', error);
    throw new Error('Failed to save custom adversary: ' + error.message);
  }
}

function deleteCustomAdversary(adversaryName) {
  try {
    const customAdversaries = loadAdversariesFromDrive('custom-adversaries');
    
    if (!customAdversaries[adversaryName]) {
      throw new Error('Adversary not found in custom adversaries');
    }
    
    delete customAdversaries[adversaryName];
    
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(customAdversaries, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName('custom-adversaries.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    }
    
    DATA_CACHE.adversaries = null;
    DATA_CACHE.lastAccessTime = 0;
    
    return { success: true };
  } catch (error) {
    console.error('Error deleting custom adversary:', error);
    throw new Error('Failed to delete custom adversary: ' + error.message);
  }
}

function saveCustomEnvironment(environment) {
  try {
    const customEnvironments = loadEnvironmentsFromDrive('custom-environments');
    customEnvironments[environment.name] = environment;
    
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(customEnvironments, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName('custom-environments.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    } else {
      encounterDataFolder.createFile('custom-environments.json', jsonContent, MimeType.PLAIN_TEXT);
    }
    
    DATA_CACHE.environments = null;
    DATA_CACHE.lastAccessTime = 0;
    
    return { success: true };
  } catch (error) {
    console.error('Error saving custom environment:', error);
    throw new Error('Failed to save custom environment: ' + error.message);
  }
}

function deleteCustomEnvironment(environmentName) {
  try {
    const customEnvironments = loadEnvironmentsFromDrive('custom-environments');
    
    if (!customEnvironments[environmentName]) {
      throw new Error('Environment not found in custom environments');
    }
    
    delete customEnvironments[environmentName];
    
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(customEnvironments, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName('custom-environments.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    }
    
    DATA_CACHE.environments = null;
    DATA_CACHE.lastAccessTime = 0;
    
    return { success: true };
  } catch (error) {
    console.error('Error deleting custom environment:', error);
    throw new Error('Failed to delete custom environment: ' + error.message);
  }
}

function showSettingsDialog() {
  loadConfiguration();
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = createSettingsDialog();
  const dialogSize = DIALOG_PRESETS.ADVERSARY_EDITOR || { WIDTH: 1000, HEIGHT: 800 };
  htmlOutput.setWidth(dialogSize.WIDTH).setHeight(dialogSize.HEIGHT);
  ui.showModalDialog(htmlOutput, 'System Settings');
}

function createSettingsDialog() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          ${getSettingsCSS()}
        </style>
      </head>
      <body>
        <div class="settings-container">
          ${generateSettingsHTML()}
        </div>
        <script>
          ${generateSettingsScript()}
        </script>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

function getSettingsCSS() {
  return `
    * { box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 20px; 
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      min-height: 100vh;
    }
    .settings-container {
      background: white;
      padding: 30px;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
      max-width: 1200px;
      margin: 0 auto;
      height: calc(100vh - 40px);
      overflow-y: auto;
    }
    .settings-header {
      text-align: center;
      margin-bottom: 30px;
      padding-bottom: 20px;
      border-bottom: 3px solid #e3f2fd;
    }
    .settings-header h2 {
      margin: 0;
      color: #1976d2;
      font-size: 24px;
      font-weight: 700;
    }
    .settings-section {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #e0e0e0;
      margin-bottom: 20px;
    }
    .settings-section h3 {
      margin: 0 0 15px 0;
      color: #1976d2;
      font-size: 18px;
      font-weight: 600;
      border-bottom: 2px solid #e3f2fd;
      padding-bottom: 8px;
    }
    .settings-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
    }
    .settings-row {
      margin-bottom: 15px;
    }
    .settings-row:last-child {
      margin-bottom: 0;
    }
    .settings-row label {
      display: block;
      font-weight: 600;
      color: #333;
      margin-bottom: 5px;
      font-size: 14px;
    }
    .settings-row input, .settings-row select {
      width: 100%;
      padding: 10px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: all 0.3s ease;
    }
    .settings-row input:focus, .settings-row select:focus {
      border-color: #1976d2;
      outline: none;
      box-shadow: 0 0 0 3px rgba(25,118,210,0.2);
    }
    .settings-row small {
      display: block;
      color: #666;
      font-size: 12px;
      margin-top: 3px;
      font-style: italic;
    }
    .checkbox-row {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 10px;
    }
    .checkbox-row input[type="checkbox"] {
      width: auto;
      transform: scale(1.2);
      accent-color: #1976d2;
    }
    .checkbox-row label {
      margin: 0;
      font-weight: normal;
    }
    .action-buttons {
      display: flex;
      gap: 15px;
      justify-content: center;
      margin-top: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .btn {
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
    }
    .btn-primary {
      background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
      color: white;
    }
    .btn-secondary {
      background: linear-gradient(135deg, #757575 0%, #616161 100%);
      color: white;
    }
    .btn-danger {
      background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);
      color: white;
    }
    .btn:disabled {
      background: #bdbdbd;
      cursor: not-allowed;
    }
    .version-info {
      text-align: center;
      font-size: 11px;
      color: #999;
      margin-top: 20px;
      font-style: italic;
    }
    .warning-banner {
      background: #fff3cd;
      border: 1px solid #ffeaa7;
      color: #856404;
      padding: 12px;
      border-radius: 8px;
      margin-bottom: 20px;
      font-size: 14px;
    }
    .warning-banner strong {
      color: #721c24;
    }
    @media (max-width: 900px) {
      .settings-grid {
        grid-template-columns: 1fr;
      }
    }
  `;
}

function generateSettingsHTML() {
  return `
    <div class="settings-header">
      <h2>System Settings</h2>
    </div>
    
    <div class="warning-banner">
      <strong>⚠️ Important:</strong> Changes will take effect after saving and may require clearing the data cache. 
      Advanced settings can affect system performance and behavior.
    </div>
    
    <div class="settings-grid">
      <div class="settings-section">
        <h3>Sheet Configuration</h3>
        <div class="settings-row">
          <label>Encounter Sheet Name:</label>
          <input type="text" id="encounter-sheet" value="">
          <small>Name for the main encounter sheets</small>
        </div>
        <div class="settings-row">
          <label>Health Tracker Sheet:</label>
          <input type="text" id="tracker-sheet" value="">
          <small>Name for health tracking sheets</small>
        </div>
        <div class="settings-row">
          <label>Adversary Data Sheet:</label>
          <input type="text" id="adversary-sheet" value="">
          <small>Name for adversary template sheets</small>
        </div>
        <div class="settings-row">
          <label>Custom Adversaries Sheet:</label>
          <input type="text" id="custom-sheet" value="">
          <small>Name for custom adversary sheets</small>
        </div>
      </div>
      
      <div class="settings-section">
        <h3>Battle Point System</h3>
        <div class="settings-row">
          <label>Base Multiplier:</label>
          <input type="number" id="base-multiplier" min="1" max="10" value="">
          <small>Multiplied by player count in base calculation</small>
        </div>
        <div class="settings-row">
          <label>Base Addition:</label>
          <input type="number" id="base-addition" min="0" max="10" value="">
          <small>Added to base calculation result</small>
        </div>
        <div class="settings-row">
          <label>Default Party Size:</label>
          <input type="number" id="default-party-size" min="1" max="8" value="">
          <small>Default player count for encounters</small>
        </div>
        <div class="settings-row">
          <label>Default Party Tier:</label>
          <select id="default-party-tier">
            <option value="1">Tier 1</option>
            <option value="2">Tier 2</option>
            <option value="3">Tier 3</option>
            <option value="4">Tier 4</option>
          </select>
          <small>Default tier selection for encounters</small>
        </div>
      </div>
    </div>
    
    <div class="settings-grid">
      <div class="settings-section">
        <h3>Battle Values</h3>
        <div class="settings-row">
          <label>Minion Value:</label>
          <input type="number" id="minion-value" min="0" max="10" value="">
          <small>Battle points for Minion type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Standard Value:</label>
          <input type="number" id="standard-value" min="0" max="10" value="">
          <small>Battle points for Standard type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Support Value:</label>
          <input type="number" id="support-value" min="0" max="10" value="">
          <small>Battle points for Support type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Social Value:</label>
          <input type="number" id="social-value" min="0" max="10" value="">
          <small>Battle points for Social type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Leader Value:</label>
          <input type="number" id="leader-value" min="0" max="10" value="">
          <small>Battle points for Leader type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Bruiser Value:</label>
          <input type="number" id="bruiser-value" min="0" max="10" value="">
          <small>Battle points for Bruiser type adversaries</small>
        </div>
        <div class="settings-row">
          <label>Solo Value:</label>
          <input type="number" id="solo-value" min="0" max="10" value="">
          <small>Battle points for Solo type adversaries</small>
        </div>
      </div>
      
      <div class="settings-section">
        <h3>GitHub Configuration</h3>
        <div class="settings-row">
          <label>Repository Base URL:</label>
          <input type="text" id="repo-base" value="">
          <small>Base URL for GitHub API repository access</small>
        </div>
        <div class="settings-row">
          <label>Adversaries Path:</label>
          <input type="text" id="adversaries-path" value="">
          <small>Path to adversaries folder in repository</small>
        </div>
        <div class="settings-row">
          <label>Environments Path:</label>
          <input type="text" id="environments-path" value="">
          <small>Path to environments folder in repository</small>
        </div>
      </div>
    </div>
    
    <div class="settings-grid">
      <div class="settings-section">
        <h3>Battle Adjustments</h3>
        <div class="settings-row">
          <label>Multiple Solos Adjustment:</label>
          <input type="number" id="solos-adjustment" min="-10" max="0" value="">
          <small>Points adjustment when 2+ solos present</small>
        </div>
        <div class="settings-row">
          <label>Min Solos for Adjustment:</label>
          <input type="number" id="min-solos" min="2" max="5" value="">
          <small>Minimum solos needed for adjustment</small>
        </div>
        <div class="settings-row">
          <label>Lower Tier Bonus:</label>
          <input type="number" id="lower-tier" min="0" max="5" value="">
          <small>Points added for lower tier adversaries</small>
        </div>
        <div class="settings-row">
          <label>No Elites Bonus:</label>
          <input type="number" id="no-elites" min="0" max="5" value="">
          <small>Points added when no elite types present</small>
        </div>
        <div class="settings-row">
          <label>High Damage Penalty:</label>
          <input type="number" id="high-damage" min="-10" max="0" value="">
          <small>Points adjustment for high damage encounters</small>
        </div>
      </div>
      
      <div class="settings-section">
        <h3>Dialog Sizes</h3>
        <div class="settings-row">
          <label>Encounter Builder Width:</label>
          <input type="number" id="encounter-width" min="800" max="1600" value="">
          <small>Width of Encounter Builder dialog</small>
        </div>
        <div class="settings-row">
          <label>Encounter Builder Height:</label>
          <input type="number" id="encounter-height" min="600" max="1000" value="">
          <small>Height of Encounter Builder dialog</small>
        </div>
        <div class="settings-row">
          <label>Adversary Editor Width:</label>
          <input type="number" id="adversary-width" min="800" max="1400" value="">
          <small>Width of Adversary Editor dialog</small>
        </div>
        <div class="settings-row">
          <label>Adversary Editor Height:</label>
          <input type="number" id="adversary-height" min="600" max="1000" value="">
          <small>Height of Adversary Editor dialog</small>
        </div>
        <div class="settings-row">
          <label>Environment Editor Width:</label>
          <input type="number" id="environment-width" min="800" max="1400" value="">
          <small>Width of Environment Editor dialog</small>
        </div>
        <div class="settings-row">
          <label>Environment Editor Height:</label>
          <input type="number" id="environment-height" min="600" max="1000" value="">
          <small>Height of Environment Editor dialog</small>
        </div>
      </div>
    </div>
    
    <div class="settings-grid">
      <div class="settings-section">
        <h3>User Interface</h3>
        <div class="settings-row">
          <label>Legacy Dialog Width:</label>
          <input type="number" id="dialog-width" min="800" max="1600" value="">
          <small>Default width for legacy dialog components</small>
        </div>
        <div class="settings-row">
          <label>Legacy Dialog Height:</label>
          <input type="number" id="dialog-height" min="600" max="1000" value="">
          <small>Default height for legacy dialog components</small>
        </div>
        <div class="settings-row">
          <label>Search Results Max Height (px):</label>
          <input type="number" id="results-height" min="300" max="800" value="">
          <small>Maximum height for search results area</small>
        </div>
        <div class="settings-row">
          <label>Search Debounce Delay (ms):</label>
          <input type="number" id="search-delay" min="100" max="1000" value="">
          <small>Delay before executing search</small>
        </div>
        <div class="settings-row">
          <label>Auto-Close Delay (ms):</label>
          <input type="number" id="auto-close" min="5000" max="30000" value="">
          <small>Time before dialogs auto-close</small>
        </div>
        <div class="settings-row">
          <label>Loading Spinner Delay (ms):</label>
          <input type="number" id="spinner-delay" min="500" max="3000" value="">
          <small>Delay before showing progress spinner</small>
        </div>
      </div>
      
      <div class="settings-section">
        <h3>Sheet Formatting</h3>
        <div class="settings-row">
          <label>Name Column Width:</label>
          <input type="number" id="name-column" min="100" max="300" value="">
          <small>Width of adversary name columns</small>
        </div>
        <div class="settings-row">
          <label>Difficulty Column Width:</label>
          <input type="number" id="difficulty-column" min="60" max="150" value="">
          <small>Width of difficulty columns</small>
        </div>
        <div class="settings-row">
          <label>Thresholds Column Width:</label>
          <input type="number" id="thresholds-column" min="80" max="150" value="">
          <small>Width of threshold columns</small>
        </div>
        <div class="settings-row">
          <label>Attack Column Width:</label>
          <input type="number" id="attack-column" min="60" max="120" value="">
          <small>Width of attack modifier columns</small>
        </div>
        <div class="settings-row">
          <label>Standard Attack Column Width:</label>
          <input type="number" id="standard-attack-column" min="200" max="400" value="">
          <small>Width of standard attack description columns</small>
        </div>
        <div class="settings-row">
          <label>Checkbox Column Width:</label>
          <input type="number" id="checkbox-column" min="30" max="80" value="">
          <small>Width of HP/Stress checkbox columns</small>
        </div>
        <div class="settings-row">
          <label>Notes Column Width:</label>
          <input type="number" id="notes-column" min="150" max="400" value="">
          <small>Width of notes columns</small>
        </div>
      </div>
    </div>
    
    <div class="settings-grid">
      <div class="settings-section">
        <h3>Performance & Caching</h3>
        <div class="settings-row">
          <label>Cache Duration (minutes):</label>
          <input type="number" id="cache-duration" min="1" max="60" value="">
          <small>How long to cache data in memory</small>
        </div>
        <div class="settings-row">
          <label>GitHub Request Delay (ms):</label>
          <input type="number" id="github-delay" min="100" max="1000" value="">
          <small>Delay between GitHub API requests</small>
        </div>
        <div class="settings-row">
          <label>Max Concurrent Requests:</label>
          <input type="number" id="max-requests" min="1" max="10" value="">
          <small>Maximum simultaneous GitHub requests</small>
        </div>
        <div class="settings-row">
          <label>Retry Attempts:</label>
          <input type="number" id="retry-attempts" min="1" max="5" value="">
          <small>Number of retry attempts for failed requests</small>
        </div>
      </div>
      
      <div class="settings-section">
        <h3>Feature Flags</h3>
        <div class="checkbox-row">
          <input type="checkbox" id="advanced-search">
          <label>Enable Advanced Search</label>
        </div>
        <div class="checkbox-row">
          <input type="checkbox" id="bulk-operations">
          <label>Enable Bulk Operations</label>
        </div>
        <div class="checkbox-row">
          <input type="checkbox" id="export-functions">
          <label>Enable Export Functions</label>
        </div>
        <div class="checkbox-row">
          <input type="checkbox" id="backup-system">
          <label>Enable Backup System</label>
        </div>
        <div class="checkbox-row">
          <input type="checkbox" id="debug-logging">
          <label>Enable Debug Logging</label>
        </div>
      </div>
    </div>
    
    <div class="action-buttons">
      <button class="btn btn-primary" onclick="saveSettings()" id="save-btn">Save Settings</button>
      <button class="btn btn-secondary" onclick="resetToDefaults()">Reset to Defaults</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </div>
    
    <div class="version-info">System Settings 0.2 - Configuration Management</div>
  `;
}

function generateSettingsScript() {
  const currentConfig = {
    CONFIG: CONFIG,
    BATTLE_VALUES: BATTLE_VALUES,
    UI_CONFIG: UI_CONFIG,
    GITHUB_CONFIG: GITHUB_CONFIG,
    BATTLE_CALCULATION: BATTLE_CALCULATION,
    SEARCH_CONFIG: SEARCH_CONFIG,
    SYSTEM_TIMINGS: SYSTEM_TIMINGS,
    CACHE_CONFIG: CACHE_CONFIG,
    GITHUB_RATE_LIMITS: GITHUB_RATE_LIMITS,
    SHEET_FORMATTING: SHEET_FORMATTING,
    DIALOG_PRESETS: DIALOG_PRESETS,
    FEATURE_FLAGS: FEATURE_FLAGS
  };
  
  return `
    const currentSettings = ${JSON.stringify(currentConfig)};
    
    function loadCurrentSettings() {
      // Sheet Configuration
      document.getElementById('encounter-sheet').value = currentSettings.CONFIG.ENCOUNTER_SHEET || '';
      document.getElementById('tracker-sheet').value = currentSettings.CONFIG.TRACKER_SHEET || '';
      document.getElementById('adversary-sheet').value = currentSettings.CONFIG.ADVERSARY_DATA_SHEET || '';
      document.getElementById('custom-sheet').value = currentSettings.CONFIG.CUSTOM_ADVERSARIES_SHEET || '';
      
      // Battle Point System
      document.getElementById('base-multiplier').value = currentSettings.BATTLE_CALCULATION.BASE_MULTIPLIER || 3;
      document.getElementById('base-addition').value = currentSettings.BATTLE_CALCULATION.BASE_ADDITION || 2;
      document.getElementById('default-party-size').value = currentSettings.BATTLE_CALCULATION.DEFAULT_PARTY_SIZE || 4;
      document.getElementById('default-party-tier').value = currentSettings.BATTLE_CALCULATION.DEFAULT_PARTY_TIER || 2;
      
      // Battle Values
      document.getElementById('minion-value').value = currentSettings.BATTLE_VALUES.Minion || 0;
      document.getElementById('standard-value').value = currentSettings.BATTLE_VALUES.Standard || 2;
      document.getElementById('support-value').value = currentSettings.BATTLE_VALUES.Support || 1;
      document.getElementById('social-value').value = currentSettings.BATTLE_VALUES.Social || 1;
      document.getElementById('leader-value').value = currentSettings.BATTLE_VALUES.Leader || 3;
      document.getElementById('bruiser-value').value = currentSettings.BATTLE_VALUES.Bruiser || 4;
      document.getElementById('solo-value').value = currentSettings.BATTLE_VALUES.Solo || 5;
      
      // GitHub Configuration
      document.getElementById('repo-base').value = currentSettings.GITHUB_CONFIG.REPO_BASE || '';
      document.getElementById('adversaries-path').value = currentSettings.GITHUB_CONFIG.ADVERSARIES_PATH || '';
      document.getElementById('environments-path').value = currentSettings.GITHUB_CONFIG.ENVIRONMENTS_PATH || '';
      
      // Battle Adjustments
      document.getElementById('solos-adjustment').value = currentSettings.BATTLE_CALCULATION.MULTIPLE_SOLOS_ADJUSTMENT || -2;
      document.getElementById('min-solos').value = currentSettings.BATTLE_CALCULATION.MIN_SOLOS_FOR_ADJUSTMENT || 2;
      document.getElementById('lower-tier').value = currentSettings.BATTLE_CALCULATION.LOWER_TIER_BONUS || 1;
      document.getElementById('no-elites').value = currentSettings.BATTLE_CALCULATION.NO_ELITES_BONUS || 1;
      document.getElementById('high-damage').value = currentSettings.BATTLE_CALCULATION.HIGH_DAMAGE_PENALTY || -2;
      
      // Dialog Sizes
      document.getElementById('encounter-width').value = currentSettings.DIALOG_PRESETS.ENCOUNTER_BUILDER?.WIDTH || 1100;
      document.getElementById('encounter-height').value = currentSettings.DIALOG_PRESETS.ENCOUNTER_BUILDER?.HEIGHT || 750;
      document.getElementById('adversary-width').value = currentSettings.DIALOG_PRESETS.ADVERSARY_EDITOR?.WIDTH || 1000;
      document.getElementById('adversary-height').value = currentSettings.DIALOG_PRESETS.ADVERSARY_EDITOR?.HEIGHT || 800;
      document.getElementById('environment-width').value = currentSettings.DIALOG_PRESETS.ENVIRONMENT_EDITOR?.WIDTH || 1000;
      document.getElementById('environment-height').value = currentSettings.DIALOG_PRESETS.ENVIRONMENT_EDITOR?.HEIGHT || 800;
      
      // User Interface
      document.getElementById('dialog-width').value = currentSettings.UI_CONFIG.DIALOG_WIDTH || 1100;
      document.getElementById('dialog-height').value = currentSettings.UI_CONFIG.DIALOG_HEIGHT || 750;
      document.getElementById('results-height').value = currentSettings.SEARCH_CONFIG.RESULTS_MAX_HEIGHT || 450;
      document.getElementById('search-delay').value = currentSettings.SEARCH_CONFIG.SEARCH_DEBOUNCE_DELAY || 300;
      document.getElementById('auto-close').value = currentSettings.SYSTEM_TIMINGS.AUTO_CLOSE_DELAY || 10000;
      document.getElementById('spinner-delay').value = currentSettings.SYSTEM_TIMINGS.PROGRESS_SPINNER_DELAY || 1000;
      
      // Sheet Formatting
      document.getElementById('name-column').value = currentSettings.SHEET_FORMATTING.NAME_COLUMN_WIDTH || 150;
      document.getElementById('difficulty-column').value = currentSettings.SHEET_FORMATTING.DIFFICULTY_COLUMN_WIDTH || 90;
      document.getElementById('thresholds-column').value = currentSettings.SHEET_FORMATTING.THRESHOLDS_COLUMN_WIDTH || 100;
      document.getElementById('attack-column').value = currentSettings.SHEET_FORMATTING.ATTACK_COLUMN_WIDTH || 80;
      document.getElementById('standard-attack-column').value = currentSettings.SHEET_FORMATTING.STANDARD_ATTACK_COLUMN_WIDTH || 250;
      document.getElementById('checkbox-column').value = currentSettings.SHEET_FORMATTING.CHECKBOX_COLUMN_WIDTH || 45;
      document.getElementById('notes-column').value = currentSettings.SHEET_FORMATTING.NOTES_COLUMN_WIDTH || 200;
      
      // Performance & Caching
      const cacheDurationMinutes = Math.round((currentSettings.CACHE_CONFIG.CACHE_DURATION || 300000) / 60000);
      document.getElementById('cache-duration').value = cacheDurationMinutes;
      document.getElementById('github-delay').value = currentSettings.GITHUB_RATE_LIMITS.REQUEST_DELAY_MS || 200;
      document.getElementById('max-requests').value = currentSettings.GITHUB_RATE_LIMITS.MAX_CONCURRENT_REQUESTS || 5;
      document.getElementById('retry-attempts').value = currentSettings.GITHUB_RATE_LIMITS.RETRY_ATTEMPTS || 3;
      
      // Feature Flags
      document.getElementById('advanced-search').checked = currentSettings.FEATURE_FLAGS.ENABLE_ADVANCED_SEARCH || false;
      document.getElementById('bulk-operations').checked = currentSettings.FEATURE_FLAGS.ENABLE_BULK_OPERATIONS || false;
      document.getElementById('export-functions').checked = currentSettings.FEATURE_FLAGS.ENABLE_EXPORT_FUNCTIONS || false;
      document.getElementById('backup-system').checked = currentSettings.FEATURE_FLAGS.ENABLE_BACKUP_SYSTEM || false;
      document.getElementById('debug-logging').checked = currentSettings.FEATURE_FLAGS.ENABLE_DEBUG_LOGGING || false;
    }
    
    function saveSettings() {
      const saveBtn = document.getElementById('save-btn');
      saveBtn.innerHTML = 'Saving...';
      saveBtn.disabled = true;
      
      const newConfig = {
        CONFIG: {
          ENCOUNTER_SHEET: document.getElementById('encounter-sheet').value.trim(),
          TRACKER_SHEET: document.getElementById('tracker-sheet').value.trim(),
          ADVERSARY_DATA_SHEET: document.getElementById('adversary-sheet').value.trim(),
          CUSTOM_ADVERSARIES_SHEET: document.getElementById('custom-sheet').value.trim()
        },
        BATTLE_VALUES: {
          Minion: parseInt(document.getElementById('minion-value').value),
          Standard: parseInt(document.getElementById('standard-value').value),
          Horde: parseInt(document.getElementById('standard-value').value), // Use Standard value for Horde
          Skulk: parseInt(document.getElementById('standard-value').value), // Use Standard value for Skulk 
          Ranged: parseInt(document.getElementById('standard-value').value), // Use Standard value for Ranged
          Support: parseInt(document.getElementById('support-value').value),
          Social: parseInt(document.getElementById('social-value').value),
          Leader: parseInt(document.getElementById('leader-value').value),
          Bruiser: parseInt(document.getElementById('bruiser-value').value),
          Solo: parseInt(document.getElementById('solo-value').value)
        },
        UI_CONFIG: {
          DIALOG_WIDTH: parseInt(document.getElementById('dialog-width').value),
          DIALOG_HEIGHT: parseInt(document.getElementById('dialog-height').value),
          RESULTS_MAX_HEIGHT: parseInt(document.getElementById('results-height').value),
          AUTO_CLOSE_DELAY: parseInt(document.getElementById('auto-close').value),
          PROGRESS_SPINNER_DELAY: parseInt(document.getElementById('spinner-delay').value)
        },
        GITHUB_CONFIG: {
          REPO_BASE: document.getElementById('repo-base').value.trim(),
          ADVERSARIES_PATH: document.getElementById('adversaries-path').value.trim(),
          ENVIRONMENTS_PATH: document.getElementById('environments-path').value.trim()
        },
        BATTLE_CALCULATION: {
          BASE_MULTIPLIER: parseInt(document.getElementById('base-multiplier').value),
          BASE_ADDITION: parseInt(document.getElementById('base-addition').value),
          DEFAULT_PARTY_SIZE: parseInt(document.getElementById('default-party-size').value),
          DEFAULT_PARTY_TIER: parseInt(document.getElementById('default-party-tier').value),
          MULTIPLE_SOLOS_ADJUSTMENT: parseInt(document.getElementById('solos-adjustment').value),
          MIN_SOLOS_FOR_ADJUSTMENT: parseInt(document.getElementById('min-solos').value),
          LOWER_TIER_BONUS: parseInt(document.getElementById('lower-tier').value),
          NO_ELITES_BONUS: parseInt(document.getElementById('no-elites').value),
          HIGH_DAMAGE_PENALTY: parseInt(document.getElementById('high-damage').value)
        },
        SEARCH_CONFIG: {
          RESULTS_MAX_HEIGHT: parseInt(document.getElementById('results-height').value),
          SEARCH_DEBOUNCE_DELAY: parseInt(document.getElementById('search-delay').value),
          MIN_SEARCH_CHARS: currentSettings.SEARCH_CONFIG.MIN_SEARCH_CHARS || 1,
          MAX_SEARCH_RESULTS: currentSettings.SEARCH_CONFIG.MAX_SEARCH_RESULTS || 100
        },
        SYSTEM_TIMINGS: {
          AUTO_CLOSE_DELAY: parseInt(document.getElementById('auto-close').value),
          PROGRESS_SPINNER_DELAY: parseInt(document.getElementById('spinner-delay').value),
          LOADING_ANIMATION_DURATION: currentSettings.SYSTEM_TIMINGS.LOADING_ANIMATION_DURATION || 2000,
          ERROR_MESSAGE_TIMEOUT: currentSettings.SYSTEM_TIMINGS.ERROR_MESSAGE_TIMEOUT || 5000
        },
        CACHE_CONFIG: {
          CACHE_DURATION: parseInt(document.getElementById('cache-duration').value) * 60000,
          MIN_CACHE_CHECK_INTERVAL: currentSettings.CACHE_CONFIG.MIN_CACHE_CHECK_INTERVAL || 1000
        },
        GITHUB_RATE_LIMITS: {
          REQUEST_DELAY_MS: parseInt(document.getElementById('github-delay').value),
          MAX_CONCURRENT_REQUESTS: parseInt(document.getElementById('max-requests').value),
          RETRY_ATTEMPTS: parseInt(document.getElementById('retry-attempts').value)
        },
        SHEET_FORMATTING: {
          NAME_COLUMN_WIDTH: parseInt(document.getElementById('name-column').value),
          DIFFICULTY_COLUMN_WIDTH: parseInt(document.getElementById('difficulty-column').value),
          THRESHOLDS_COLUMN_WIDTH: parseInt(document.getElementById('thresholds-column').value),
          ATTACK_COLUMN_WIDTH: parseInt(document.getElementById('attack-column').value),
          STANDARD_ATTACK_COLUMN_WIDTH: parseInt(document.getElementById('standard-attack-column').value),
          CHECKBOX_COLUMN_WIDTH: parseInt(document.getElementById('checkbox-column').value),
          NOTES_COLUMN_WIDTH: parseInt(document.getElementById('notes-column').value),
          MAX_COLUMNS_IN_SHEET: currentSettings.SHEET_FORMATTING.MAX_COLUMNS_IN_SHEET || 12
        },
        DIALOG_PRESETS: {
          ENCOUNTER_BUILDER: {
            WIDTH: parseInt(document.getElementById('encounter-width').value),
            HEIGHT: parseInt(document.getElementById('encounter-height').value)
          },
          ADVERSARY_EDITOR: {
            WIDTH: parseInt(document.getElementById('adversary-width').value),
            HEIGHT: parseInt(document.getElementById('adversary-height').value)
          },
          ENVIRONMENT_EDITOR: {
            WIDTH: parseInt(document.getElementById('environment-width').value),
            HEIGHT: parseInt(document.getElementById('environment-height').value)
          },
          PROGRESS_DIALOG: currentSettings.DIALOG_PRESETS.PROGRESS_DIALOG || { WIDTH: 500, HEIGHT: 300 },
          COMPLETION_DIALOG: currentSettings.DIALOG_PRESETS.COMPLETION_DIALOG || { WIDTH: 600, HEIGHT: 400 }
        },
        FEATURE_FLAGS: {
          ENABLE_ADVANCED_SEARCH: document.getElementById('advanced-search').checked,
          ENABLE_BULK_OPERATIONS: document.getElementById('bulk-operations').checked,
          ENABLE_EXPORT_FUNCTIONS: document.getElementById('export-functions').checked,
          ENABLE_BACKUP_SYSTEM: document.getElementById('backup-system').checked,
          ENABLE_DEBUG_LOGGING: document.getElementById('debug-logging').checked
        }
      };
      
      google.script.run
        .withSuccessHandler((result) => {
          if (result.success) {
            alert('Settings saved successfully! Changes will take effect after clearing the data cache.');
            google.script.host.close();
          } else {
            alert('Error saving settings: ' + result.error);
            saveBtn.innerHTML = 'Save Settings';
            saveBtn.disabled = false;
          }
        })
        .withFailureHandler((error) => {
          alert('Error saving settings: ' + error.message);
          saveBtn.innerHTML = 'Save Settings';
          saveBtn.disabled = false;
        })
        .saveSystemSettings(newConfig);
    }
    
    function resetToDefaults() {
      if (confirm('Are you sure you want to reset all settings to default values? This action cannot be undone.')) {
        google.script.run
          .withSuccessHandler(() => {
            alert('Settings reset to defaults successfully! Please refresh and clear cache.');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error resetting settings: ' + error.message);
          })
          .resetSettingsToDefaults();
      }
    }
    
    // Load current settings on page load
    loadCurrentSettings();
  `;
}

function saveSystemSettings(newConfig) {
  try {
    // Merge with existing configuration sections not exposed in UI
    const fullConfig = {
      ...newConfig,
      DAMAGE_TYPES: DAMAGE_TYPES,
      FORM_VALIDATION: FORM_VALIDATION
    };
    
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const jsonContent = JSON.stringify(fullConfig, null, 2);
    
    const existingFiles = encounterDataFolder.getFilesByName('config.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setContent(jsonContent);
    } else {
      encounterDataFolder.createFile('config.json', jsonContent, MimeType.PLAIN_TEXT);
    }
    
    // Reload configuration
    assignConfigurationValues(fullConfig);
    
    return { success: true };
  } catch (error) {
    console.error('Error saving system settings:', error);
    return { success: false, error: error.message };
  }
}

function resetSettingsToDefaults() {
  try {
    // Delete existing config file and recreate with defaults
    const encounterDataFolder = getOrCreateEncounterDataFolder();
    const existingFiles = encounterDataFolder.getFilesByName('config.json');
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      existingFile.setTrashed(true);
    }
    
    // Recreate default configuration
    createDefaultConfiguration();
    
    return { success: true };
  } catch (error) {
    console.error('Error resetting settings:', error);
    throw new Error('Failed to reset settings: ' + error.message);
  }
}
