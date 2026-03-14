const meetings = [
  {
    id: "runion-12-03",
    title: "Réunion chantier - 12 mars",
    comptesRendus: [
      {
        ref: "CR-101",
        reminder: { status: "open", assignee: "Alice" },
        entreprise: "BatiPlus",
        zone: "Nord",
        lot: "CVC",
        dueDate: "2026-03-18"
      },
      {
        ref: "CR-102",
        reminder: { status: "open", assignee: "Yanis" },
        entreprise: "MecaSud",
        zone: "Nord",
        lot: "Électricité",
        dueDate: "2026-03-20"
      },
      {
        ref: "CR-103",
        reminder: { status: "closed", assignee: "Alice" },
        entreprise: "BatiPlus",
        zone: "Sud",
        lot: "Gros oeuvre",
        dueDate: "2026-03-22"
      },
      {
        ref: "CR-104",
        reminder: { status: "open", assignee: "Sami" },
        entreprise: "TerraPro",
        zone: "Sud",
        lot: "Voirie",
        dueDate: "2026-03-25"
      }
    ]
  },
  {
    id: "runion-19-03",
    title: "Réunion synthèse - 19 mars",
    comptesRendus: [
      {
        ref: "CR-205",
        reminder: { status: "open", assignee: "Alice" },
        entreprise: "MecaSud",
        zone: "Est",
        lot: "Plomberie",
        dueDate: "2026-03-29"
      },
      {
        ref: "CR-206",
        reminder: { status: "open", assignee: "Nina" },
        entreprise: "BatiPlus",
        zone: "Est",
        lot: "CVC",
        dueDate: "2026-03-31"
      },
      {
        ref: "CR-207",
        reminder: { status: "closed", assignee: "Sami" },
        entreprise: "TerraPro",
        zone: "Ouest",
        lot: "Façade",
        dueDate: "2026-04-02"
      }
    ]
  }
];

const ui = {
  meetingSelect: document.getElementById("meeting-select"),
  zoneFilter: document.getElementById("zone-filter"),
  lotFilter: document.getElementById("lot-filter"),
  openReminders: document.getElementById("open-reminders"),
  assigneeBreakdown: document.getElementById("assignee-breakdown"),
  companyBreakdown: document.getElementById("company-breakdown"),
  calendar: document.getElementById("calendar"),
  calendarSummary: document.getElementById("calendar-summary")
};

function countBy(items, getKey) {
  return items.reduce((map, item) => {
    const key = getKey(item);
    map.set(key, (map.get(key) || 0) + 1);
    return map;
  }, new Map());
}

function toSortedEntries(map) {
  return [...map.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
}

function populateMeetings() {
  meetings.forEach((meeting) => {
    const option = document.createElement("option");
    option.value = meeting.id;
    option.textContent = meeting.title;
    ui.meetingSelect.append(option);
  });
}

function populateFilters(compteRendus) {
  const zones = [...new Set(compteRendus.map((item) => item.zone))].sort();
  const lots = [...new Set(compteRendus.map((item) => item.lot))].sort();

  for (const select of [ui.zoneFilter, ui.lotFilter]) {
    while (select.options.length > 1) {
      select.remove(1);
    }
  }

  zones.forEach((zone) => ui.zoneFilter.add(new Option(zone, zone)));
  lots.forEach((lot) => ui.lotFilter.add(new Option(lot, lot)));
}

function getSelectedMeeting() {
  return meetings.find((meeting) => meeting.id === ui.meetingSelect.value);
}

function filterRendus(compteRendus) {
  return compteRendus.filter((item) => {
    const zoneMatch = ui.zoneFilter.value === "all" || item.zone === ui.zoneFilter.value;
    const lotMatch = ui.lotFilter.value === "all" || item.lot === ui.lotFilter.value;
    return zoneMatch && lotMatch;
  });
}

function renderList(element, entries, suffix = "") {
  element.innerHTML = "";
  if (!entries.length) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = "Aucune donnée";
    element.append(li);
    return;
  }

  entries.forEach(([label, count]) => {
    const li = document.createElement("li");
    li.textContent = `${label}: ${count}${suffix}`;
    element.append(li);
  });
}

function renderCalendar(items) {
  ui.calendar.innerHTML = "";
  if (!items.length) {
    ui.calendar.innerHTML = '<p class="empty">Aucun rendu pour ce filtre.</p>';
    return;
  }

  const sorted = [...items].sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));

  sorted.forEach((item) => {
    const article = document.createElement("article");
    article.className = "calendar-item";

    const dueDate = new Date(item.dueDate).toLocaleDateString("fr-FR", {
      weekday: "short",
      day: "2-digit",
      month: "short"
    });

    article.innerHTML = `
      <div>
        <strong>${item.ref}</strong>
        <div class="calendar-meta">${item.entreprise} · attribué à ${item.reminder.assignee}</div>
      </div>
      <div class="calendar-chips">
        <span class="chip">${dueDate}</span>
        <span class="chip">Zone ${item.zone}</span>
        <span class="chip">Lot ${item.lot}</span>
      </div>
    `;

    ui.calendar.append(article);
  });
}

function refresh() {
  const meeting = getSelectedMeeting();
  const filteredRendus = filterRendus(meeting.comptesRendus);

  const openReminders = filteredRendus.filter((item) => item.reminder.status === "open");
  ui.openReminders.textContent = String(openReminders.length);

  renderList(
    ui.assigneeBreakdown,
    toSortedEntries(countBy(openReminders, (item) => item.reminder.assignee))
  );

  renderList(
    ui.companyBreakdown,
    toSortedEntries(countBy(filteredRendus, (item) => item.entreprise))
  );

  renderCalendar(filteredRendus);
  ui.calendarSummary.textContent = `${filteredRendus.length} rendu(x) affiché(s)`;
}

ui.meetingSelect.addEventListener("change", () => {
  const meeting = getSelectedMeeting();
  populateFilters(meeting.comptesRendus);
  ui.zoneFilter.value = "all";
  ui.lotFilter.value = "all";
  refresh();
});

ui.zoneFilter.addEventListener("change", refresh);
ui.lotFilter.addEventListener("change", refresh);

populateMeetings();
ui.meetingSelect.value = meetings[0].id;
populateFilters(meetings[0].comptesRendus);
refresh();
