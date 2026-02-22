const CALENDARS_STORAGE_KEY = "justcal-calendars";
const DEFAULT_CALENDAR_LABEL = "Energy Tracker";
const DEFAULT_CALENDAR_ID = "energy-tracker";
const DEFAULT_CALENDAR_COLOR = "blue";
const DEFAULT_NEW_CALENDAR_COLOR = "gray";
const SUPPORTED_CALENDAR_TYPE = "signal-3";
const CALENDAR_BUTTON_LABEL = "Open calendars";
const CALENDAR_CLOSE_LABEL = "Close calendars";
const CALENDAR_COLOR_HEX_BY_KEY = Object.freeze({
  gray: "#9ca3af",
  red: "#ef4444",
  orange: "#f97316",
  yellow: "#facc15",
  cyan: "#22d3ee",
  blue: "#3b82f6",
});

function getDefaultCalendar() {
  return {
    id: DEFAULT_CALENDAR_ID,
    name: DEFAULT_CALENDAR_LABEL,
    type: SUPPORTED_CALENDAR_TYPE,
    color: DEFAULT_CALENDAR_COLOR,
  };
}

function isObjectLike(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function hasOwnProperty(target, key) {
  return Object.prototype.hasOwnProperty.call(target, key);
}

function sanitizeCalendarName(rawName) {
  return String(rawName ?? "").replace(/\s+/g, " ").trim();
}

function normalizeCalendarColor(colorKey, fallbackColor = DEFAULT_NEW_CALENDAR_COLOR) {
  if (typeof colorKey === "string" && hasOwnProperty(CALENDAR_COLOR_HEX_BY_KEY, colorKey)) {
    return colorKey;
  }
  return fallbackColor;
}

function resolveCalendarColorHex(colorKey, fallbackColor = DEFAULT_CALENDAR_COLOR) {
  const normalizedColor = normalizeCalendarColor(colorKey, fallbackColor);
  return CALENDAR_COLOR_HEX_BY_KEY[normalizedColor];
}

function slugifyCalendarName(name) {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);
}

function createUniqueCalendarId(name, usedIds) {
  const baseId = slugifyCalendarName(name) || "calendar";
  let candidateId = baseId;
  let suffix = 2;
  while (usedIds.has(candidateId)) {
    candidateId = `${baseId}-${suffix}`;
    suffix += 1;
  }
  return candidateId;
}

function normalizeStoredCalendar(rawCalendar, index, usedIds) {
  if (!isObjectLike(rawCalendar)) return null;

  const fallbackName = `Calendar ${index + 1}`;
  const normalizedName = sanitizeCalendarName(rawCalendar.name) || fallbackName;
  const rawId = typeof rawCalendar.id === "string" ? rawCalendar.id.trim() : "";
  const normalizedId =
    rawId && !usedIds.has(rawId) ? rawId : createUniqueCalendarId(normalizedName, usedIds);

  usedIds.add(normalizedId);

  return {
    id: normalizedId,
    name: normalizedName,
    type: SUPPORTED_CALENDAR_TYPE,
    color: normalizeCalendarColor(rawCalendar.color),
  };
}

function loadCalendarsState() {
  const fallbackCalendar = getDefaultCalendar();
  const fallbackState = {
    activeCalendarId: fallbackCalendar.id,
    calendars: [fallbackCalendar],
  };

  try {
    const rawStoredValue = localStorage.getItem(CALENDARS_STORAGE_KEY);
    if (rawStoredValue === null) {
      return fallbackState;
    }

    const parsedState = JSON.parse(rawStoredValue);
    if (!isObjectLike(parsedState)) {
      return fallbackState;
    }

    const storedCalendars = Array.isArray(parsedState.calendars) ? parsedState.calendars : [];
    const usedIds = new Set();
    const normalizedCalendars = storedCalendars
      .map((calendar, index) => normalizeStoredCalendar(calendar, index, usedIds))
      .filter(Boolean);

    const calendars = normalizedCalendars.length > 0 ? normalizedCalendars : [fallbackCalendar];
    const activeCalendarIdRaw =
      typeof parsedState.activeCalendarId === "string"
        ? parsedState.activeCalendarId.trim()
        : "";
    const hasStoredActiveCalendar = calendars.some(
      (calendar) => calendar.id === activeCalendarIdRaw,
    );

    return {
      calendars,
      activeCalendarId: hasStoredActiveCalendar ? activeCalendarIdRaw : calendars[0].id,
    };
  } catch {
    return fallbackState;
  }
}

function saveCalendarsState({ calendars, activeCalendarId }) {
  try {
    localStorage.setItem(
      CALENDARS_STORAGE_KEY,
      JSON.stringify({
        version: 1,
        activeCalendarId,
        calendars,
      }),
    );
  } catch {
    // Ignore storage errors and keep behavior in-memory.
  }
}

function createCalendarOptionElement(calendar, isActive) {
  const optionButton = document.createElement("button");
  optionButton.type = "button";
  optionButton.className = "calendar-option calendar-option-main";
  optionButton.classList.toggle("is-active", isActive);
  optionButton.dataset.calendarType = calendar.type;
  optionButton.dataset.calendarId = calendar.id;
  optionButton.setAttribute("aria-label", calendar.name);
  optionButton.setAttribute("aria-pressed", String(isActive));

  const left = document.createElement("span");
  left.className = "calendar-option-left";

  const dot = document.createElement("span");
  dot.className = "calendar-option-dot";
  dot.setAttribute("aria-hidden", "true");
  dot.style.setProperty("--calendar-dot-color", resolveCalendarColorHex(calendar.color));

  const label = document.createElement("span");
  label.className = "calendar-option-label";
  label.textContent = calendar.name;

  left.append(dot, label);

  const check = document.createElement("span");
  check.className = "calendar-option-check";
  check.setAttribute("aria-hidden", "true");
  check.textContent = "✓";

  optionButton.append(left, check);
  return optionButton;
}

function setCalendarButtonLabel(button, activeCalendar) {
  if (!button) return;

  const nextLabel =
    sanitizeCalendarName(activeCalendar?.name) || DEFAULT_CALENDAR_LABEL;
  const nextDotColor = resolveCalendarColorHex(
    activeCalendar?.color,
    DEFAULT_CALENDAR_COLOR,
  );

  const existingDot = button.querySelector(".calendar-current-dot");
  const existingName = button.querySelector(".calendar-current-name");
  if (existingDot && existingName) {
    existingName.textContent = nextLabel;
    existingDot.style.setProperty("--calendar-dot-color", nextDotColor);
    return;
  }

  button.textContent = "";
  const dot = document.createElement("span");
  dot.className = "calendar-current-dot";
  dot.setAttribute("aria-hidden", "true");
  dot.style.setProperty("--calendar-dot-color", nextDotColor);

  const name = document.createElement("span");
  name.className = "calendar-current-name";
  name.textContent = nextLabel;

  button.append(dot, name);
}

function setCalendarSwitcherExpanded({ switcher, button, activeCalendar, isExpanded }) {
  const activeCalendarLabel =
    sanitizeCalendarName(activeCalendar?.name) || DEFAULT_CALENDAR_LABEL;
  setCalendarButtonLabel(button, activeCalendar);
  switcher.classList.toggle("is-expanded", isExpanded);
  button.setAttribute("aria-expanded", String(isExpanded));
  button.setAttribute(
    "aria-label",
    isExpanded
      ? CALENDAR_CLOSE_LABEL
      : `${CALENDAR_BUTTON_LABEL} (${activeCalendarLabel})`,
  );
}

function setAddCalendarEditorExpanded({
  switcher,
  addShell,
  addEditor,
  addNameInput,
  addTypeSelect,
  isExpanded,
} = {}) {
  if (!switcher || !addShell || !addEditor) return;

  switcher.classList.toggle("is-adding", isExpanded);
  addShell.classList.toggle("is-editing", isExpanded);
  addEditor.setAttribute("aria-hidden", String(!isExpanded));

  if (!isExpanded) {
    if (addNameInput) {
      addNameInput.value = "";
    }
    if (addTypeSelect) {
      addTypeSelect.value = SUPPORTED_CALENDAR_TYPE;
    }
  }
}

export function setupCalendarSwitcher(button, { onActiveCalendarChange } = {}) {
  const switcher = document.getElementById("calendar-switcher");
  const calendarList = document.getElementById("calendar-list");
  const addShell = document.getElementById("calendar-add-shell");
  const addTrigger = document.getElementById("calendar-add-trigger");
  const addEditor = document.getElementById("calendar-add-editor");
  const addSubmitButton = document.getElementById("calendar-add-submit");
  const addCancelButton = document.getElementById("calendar-add-cancel");
  const addNameInput = document.getElementById("new-calendar-name");
  const addTypeSelect = document.getElementById("new-calendar-type");
  const addColorOptions = document.getElementById("new-calendar-color");
  const addColorButtons = addColorOptions
    ? [...addColorOptions.querySelectorAll(".calendar-color-option")]
    : [];

  if (!switcher || !button || !calendarList) {
    return;
  }

  const flashMissingName = () => {
    if (!addNameInput) return;
    addNameInput.classList.remove("is-error-flash");
    // Force reflow so repeated clicks replay the animation.
    void addNameInput.offsetWidth;
    addNameInput.classList.add("is-error-flash");
  };

  const setActiveColor = (nextButton) => {
    if (!nextButton) return;
    addColorButtons.forEach((candidateButton) => {
      const isActive = candidateButton === nextButton;
      candidateButton.classList.toggle("is-active", isActive);
      candidateButton.setAttribute("aria-pressed", String(isActive));
    });
  };

  const defaultColorButton =
    addColorButtons.find((candidateButton) => {
      return candidateButton.dataset.color === DEFAULT_NEW_CALENDAR_COLOR;
    }) || addColorButtons[0];

  const resetAddColor = () => {
    setActiveColor(defaultColorButton);
  };

  const state = loadCalendarsState();
  let calendars = state.calendars;
  let activeCalendarId = state.activeCalendarId;

  const resolveActiveCalendar = () => {
    return (
      calendars.find((calendar) => calendar.id === activeCalendarId) || calendars[0] || null
    );
  };

  const notifyActiveCalendarChange = () => {
    if (typeof onActiveCalendarChange !== "function") return;
    const activeCalendar = resolveActiveCalendar();
    if (activeCalendar) {
      onActiveCalendarChange(activeCalendar);
    }
  };

  const renderCalendarList = () => {
    const fragment = document.createDocumentFragment();
    calendars.forEach((calendar) => {
      const isActive = calendar.id === activeCalendarId;
      fragment.appendChild(createCalendarOptionElement(calendar, isActive));
    });
    calendarList.replaceChildren(fragment);
  };

  const syncSwitcherButton = () => {
    const activeCalendar = resolveActiveCalendar();
    setCalendarSwitcherExpanded({
      switcher,
      button,
      activeCalendar,
      isExpanded: switcher.classList.contains("is-expanded"),
    });
  };

  const persistCalendarState = () => {
    saveCalendarsState({
      calendars,
      activeCalendarId,
    });
  };

  const syncCalendarUi = () => {
    renderCalendarList();
    syncSwitcherButton();
  };

  const setActiveCalendarId = (nextCalendarId) => {
    if (!nextCalendarId) return false;
    const hasCalendar = calendars.some((calendar) => calendar.id === nextCalendarId);
    if (!hasCalendar || nextCalendarId === activeCalendarId) {
      return false;
    }

    activeCalendarId = nextCalendarId;
    persistCalendarState();
    syncCalendarUi();
    notifyActiveCalendarChange();
    return true;
  };

  const resetAddEditor = () => {
    setAddCalendarEditorExpanded({
      switcher,
      addShell,
      addEditor,
      addNameInput,
      addTypeSelect,
      isExpanded: false,
    });
    resetAddColor();
  };

  addColorButtons.forEach((colorButton) => {
    colorButton.addEventListener("click", () => {
      setActiveColor(colorButton);
    });
  });

  resetAddEditor();
  syncCalendarUi();
  saveCalendarsState({
    calendars,
    activeCalendarId,
  });
  notifyActiveCalendarChange();

  calendarList.addEventListener("click", (event) => {
    const optionButton = event.target.closest("button.calendar-option[data-calendar-id]");
    if (!optionButton || !calendarList.contains(optionButton)) {
      return;
    }

    const nextCalendarId = optionButton.dataset.calendarId || "";
    setActiveCalendarId(nextCalendarId);
    resetAddEditor();
    setCalendarSwitcherExpanded({
      switcher,
      button,
      activeCalendar: resolveActiveCalendar(),
      isExpanded: false,
    });
  });

  if (addTrigger && addShell && addEditor) {
    addTrigger.addEventListener("click", () => {
      setAddCalendarEditorExpanded({
        switcher,
        addShell,
        addEditor,
        addNameInput,
        addTypeSelect,
        isExpanded: true,
      });
      addNameInput?.focus();
    });
  }

  if (addCancelButton) {
    addCancelButton.addEventListener("click", () => {
      resetAddEditor();
      addTrigger?.focus();
    });
  }

  addNameInput?.addEventListener("input", () => {
    addNameInput.classList.remove("is-error-flash");
  });

  addNameInput?.addEventListener("animationend", (event) => {
    if (event.animationName !== "calendar-add-input-error-flash") {
      return;
    }
    addNameInput.classList.remove("is-error-flash");
  });

  if (addSubmitButton) {
    addSubmitButton.addEventListener("click", () => {
      const nextName = sanitizeCalendarName(addNameInput?.value);
      if (!nextName) {
        flashMissingName();
        addNameInput?.focus();
        return;
      }

      const usedIds = new Set(calendars.map((calendar) => calendar.id));
      const selectedColorButton = addColorButtons.find((candidateButton) => {
        return candidateButton.classList.contains("is-active");
      });

      const nextCalendar = {
        id: createUniqueCalendarId(nextName, usedIds),
        name: nextName,
        type:
          addTypeSelect?.value === SUPPORTED_CALENDAR_TYPE
            ? SUPPORTED_CALENDAR_TYPE
            : SUPPORTED_CALENDAR_TYPE,
        color: normalizeCalendarColor(selectedColorButton?.dataset.color),
      };

      calendars = [...calendars, nextCalendar];
      activeCalendarId = nextCalendar.id;

      persistCalendarState();
      syncCalendarUi();
      notifyActiveCalendarChange();

      resetAddEditor();
      setCalendarSwitcherExpanded({
        switcher,
        button,
        activeCalendar: resolveActiveCalendar(),
        isExpanded: false,
      });
      button.focus();
    });
  }

  button.addEventListener("click", () => {
    const isExpanded = switcher.classList.contains("is-expanded");
    const nextExpanded = !isExpanded;
    if (!nextExpanded) {
      resetAddEditor();
    }
    setCalendarSwitcherExpanded({
      switcher,
      button,
      activeCalendar: resolveActiveCalendar(),
      isExpanded: nextExpanded,
    });
  });

  document.addEventListener("click", (event) => {
    if (!switcher.classList.contains("is-expanded")) {
      return;
    }
    if (switcher.contains(event.target)) {
      return;
    }
    resetAddEditor();
    setCalendarSwitcherExpanded({
      switcher,
      button,
      activeCalendar: resolveActiveCalendar(),
      isExpanded: false,
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") {
      return;
    }
    if (!switcher.classList.contains("is-expanded")) {
      return;
    }
    resetAddEditor();
    setCalendarSwitcherExpanded({
      switcher,
      button,
      activeCalendar: resolveActiveCalendar(),
      isExpanded: false,
    });
    button.focus();
  });
}
