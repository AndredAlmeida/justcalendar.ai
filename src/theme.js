const THEME_STORAGE_KEY = "justcal-theme";
const LOCAL_CALENDAR_STORAGE_CHANGED_EVENT = "justcal:local-calendar-storage-changed";
const THEME_BUTTON_LABEL = "Open Themes";
const THEME_CLOSE_LABEL = "Close themes";
const TOKYO_NIGHT_STORM_THEME = "tokyo-night-storm";
const DEFAULT_THEME = TOKYO_NIGHT_STORM_THEME;
const SOLARIZED_DARK_THEME = "solarized-dark";
const SOLARIZED_LIGHT_THEME = "solarized-light";
const RED_THEME = "red";
const LEGACY_ABYSS_THEME = "abyss";
const SUPPORTED_THEMES = [
  "light",
  "dark",
  RED_THEME,
  TOKYO_NIGHT_STORM_THEME,
  SOLARIZED_DARK_THEME,
  SOLARIZED_LIGHT_THEME,
];
const DARK_STYLE_THEMES = ["dark", RED_THEME, TOKYO_NIGHT_STORM_THEME, SOLARIZED_DARK_THEME];
const CUSTOM_THEME_CLASSES = [
  RED_THEME,
  TOKYO_NIGHT_STORM_THEME,
  SOLARIZED_DARK_THEME,
  SOLARIZED_LIGHT_THEME,
  LEGACY_ABYSS_THEME,
];
const THEME_COLORS = {
  light: "#f8fafc",
  dark: "#020617",
  [RED_THEME]: "#2b050a",
  [TOKYO_NIGHT_STORM_THEME]: "#0f1224",
  [SOLARIZED_DARK_THEME]: "#002b36",
  [SOLARIZED_LIGHT_THEME]: "#fdf6e3",
};

function getStoredTheme() {
  try {
    return localStorage.getItem(THEME_STORAGE_KEY);
  } catch {
    return null;
  }
}

function setStoredTheme(theme) {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
    return true;
  } catch {
    // Ignore storage failures and keep toggle behavior in-memory.
    return false;
  }
}

function normalizeTheme(theme) {
  if (theme === LEGACY_ABYSS_THEME) {
    return SOLARIZED_DARK_THEME;
  }
  if (SUPPORTED_THEMES.includes(theme)) {
    return theme;
  }
  return null;
}

function applyTheme({ theme, root, themeColorMeta }) {
  const isDarkStyle = DARK_STYLE_THEMES.includes(theme);
  root.classList.toggle("dark", isDarkStyle);
  document.body.classList.toggle("dark", isDarkStyle);
  CUSTOM_THEME_CLASSES.forEach((themeClass) => {
    const isActiveThemeClass = theme === themeClass;
    root.classList.toggle(themeClass, isActiveThemeClass);
    document.body.classList.toggle(themeClass, isActiveThemeClass);
  });

  if (themeColorMeta) {
    themeColorMeta.setAttribute("content", THEME_COLORS[theme] || THEME_COLORS.dark);
  }
}

function setThemeSwitcherExpanded({ switcher, button, isExpanded }) {
  switcher.classList.toggle("is-expanded", isExpanded);
  button.setAttribute("aria-expanded", String(isExpanded));
  button.setAttribute("aria-label", isExpanded ? THEME_CLOSE_LABEL : THEME_BUTTON_LABEL);
  button.setAttribute("data-tooltip", isExpanded ? THEME_CLOSE_LABEL : THEME_BUTTON_LABEL);
  button.removeAttribute("title");
}

function syncThemeOptionState({ optionButtons, theme }) {
  optionButtons.forEach((optionButton) => {
    const optionTheme = optionButton.dataset.themeOption;
    const isThemeButton = SUPPORTED_THEMES.includes(optionTheme);
    const isActive = isThemeButton && optionTheme === theme;
    optionButton.classList.toggle("is-active", isActive);
    if (isThemeButton) {
      optionButton.setAttribute("aria-pressed", String(isActive));
    }
  });
}

export function setupThemeToggle(button) {
  const switcher = document.getElementById("theme-switcher");
  const optionButtons = switcher
    ? Array.from(switcher.querySelectorAll(".theme-option"))
    : [];

  if (!switcher) {
    return;
  }

  const root = document.documentElement;
  const themeColorMeta = document.querySelector('meta[name="theme-color"]');

  const savedTheme = getStoredTheme();
  const initialTheme = normalizeTheme(savedTheme) || DEFAULT_THEME;
  let currentTheme = initialTheme;

  applyTheme({ theme: initialTheme, root, themeColorMeta });
  syncThemeOptionState({ optionButtons, theme: initialTheme });
  setThemeSwitcherExpanded({ switcher, button, isExpanded: false });

  const applyThemeSelection = (nextTheme, { persist = true } = {}) => {
    const normalizedTheme = normalizeTheme(nextTheme);
    if (!normalizedTheme) {
      return false;
    }
    applyTheme({ theme: normalizedTheme, root, themeColorMeta });
    syncThemeOptionState({ optionButtons, theme: normalizedTheme });
    if (persist) {
      const persisted = setStoredTheme(normalizedTheme);
      if (
        persisted &&
        typeof window !== "undefined" &&
        typeof window.dispatchEvent === "function"
      ) {
        window.dispatchEvent(
          new CustomEvent(LOCAL_CALENDAR_STORAGE_CHANGED_EVENT, {
            detail: {
              key: THEME_STORAGE_KEY,
            },
          }),
        );
      }
    }
    currentTheme = normalizedTheme;
    return true;
  };

  optionButtons.forEach((optionButton) => {
    optionButton.addEventListener("click", () => {
      const nextTheme = optionButton.dataset.themeOption;
      if (!SUPPORTED_THEMES.includes(nextTheme)) {
        return;
      }
      applyThemeSelection(nextTheme);
    });
  });

  button.addEventListener("click", () => {
    const isExpanded = switcher.classList.contains("is-expanded");
    setThemeSwitcherExpanded({ switcher, button, isExpanded: !isExpanded });
  });

  document.addEventListener("click", (event) => {
    if (!switcher.classList.contains("is-expanded")) {
      return;
    }
    if (switcher.contains(event.target)) {
      return;
    }
    setThemeSwitcherExpanded({ switcher, button, isExpanded: false });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") {
      return;
    }
    if (!switcher.classList.contains("is-expanded")) {
      return;
    }
    setThemeSwitcherExpanded({ switcher, button, isExpanded: false });
    button.focus();
  });

  return {
    applyTheme: (nextTheme, options = {}) => {
      return applyThemeSelection(nextTheme, options);
    },
    syncFromStorage: () => {
      const storedTheme = getStoredTheme();
      const nextTheme = normalizeTheme(storedTheme) || DEFAULT_THEME;
      applyThemeSelection(nextTheme, { persist: false });
      return nextTheme;
    },
    getTheme: () => currentTheme,
  };
}
