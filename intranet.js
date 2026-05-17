// ============================================================
// HUARD GROUPE — Intranet shared shell (topbar + sidebar + footer)
// ============================================================

// ---------- Icons (Lucide-style strokes) ----------
const ICON = {
  home:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9.5 12 2l9 7.5"/><path d="M5 9v11a1 1 0 0 0 1 1h4v-7h4v7h4a1 1 0 0 0 1-1V9"/></svg>',
  users:       '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
  folder:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h5l2 3h9a1 1 0 0 1 1 1v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1Z"/></svg>',
  briefcase:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M8 7V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>',
  building:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/></svg>',
  cart:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.7 13.4a2 2 0 0 0 2 1.6h9.7a2 2 0 0 0 2-1.6L23 6H6"/></svg>',
  wrench:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.7 2.7-2.6-2.6 2.7-2.7Z"/></svg>',
  laptop:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M2 20h20"/></svg>',
  clock:       '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
  calendar:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
  euro:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 7a6 6 0 1 0 0 10"/><path d="M7 10h8M7 14h8"/></svg>',
  fileText:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>',
  tool:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.7 6.3a4 4 0 0 0-5.4 5.4l-7.3 7.3 3 3 7.3-7.3a4 4 0 0 0 5.4-5.4l-2.7 2.7-2.6-2.6 2.7-2.7Z"/></svg>',
  package:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16.5 9.4 7.5 4.21"/><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>',
  phone:       '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.37 1.9.72 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.35 1.85.59 2.81.72A2 2 0 0 1 22 16.92Z"/></svg>',
  ticket:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 9a3 3 0 0 1 0 6v2a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-2a3 3 0 0 1 0-6V7a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v2Z"/><path d="M13 5v2M13 17v2M13 11v2"/></svg>',
  search:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
  bell:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/></svg>',
  menu:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>',
  settings:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06A2 2 0 1 1 4.27 16.97l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06A2 2 0 1 1 7.04 4.27l.06.06a1.65 1.65 0 0 0 1.82.33h.01a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82v.01a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1Z"/></svg>',
  logout:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>',
  trendingUp:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>',
  alertCircle: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
  checkCircle: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
  smartphone:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="2" width="14" height="20" rx="2"/><line x1="12" y1="18" x2="12.01" y2="18"/></svg>',
  shield:      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"/></svg>',
  zap:         '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>',
  thermometer: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 14.76V3.5a2.5 2.5 0 0 0-5 0v11.26a4.5 4.5 0 1 0 5 0Z"/></svg>',
  wifi:        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.55a11 11 0 0 1 14.08 0"/><path d="M1.42 9a16 16 0 0 1 21.16 0"/><path d="M8.53 16.11a6 6 0 0 1 6.95 0"/><line x1="12" y1="20" x2="12.01" y2="20"/></svg>',
};

// ---------- HUARD logo (logomark + wordmark) ----------
const HUARD_LOGOMARK = `
<svg viewBox="0 0 36 36" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <rect width="36" height="36" rx="8" fill="#1F2F4D"/>
  <rect x="9" y="9" width="4" height="18" fill="#FFFFFF"/>
  <rect x="23" y="9" width="4" height="18" fill="#FFFFFF"/>
  <rect x="13" y="16" width="10" height="4" fill="#2563EB"/>
</svg>`;

// ---------- Navigation ----------
const NAV_ITEMS = [
  { id: 'accueil',     label: 'Tableau de bord', href: 'intranet.html',         icon: 'home' },
  { id: 'annuaire',    label: 'Annuaire',        href: 'annuaire.html',         icon: 'phone' },
  { id: 'rh',          label: 'Ressources Humaines', href: 'rh.html',           icon: 'users' },
  { id: 'chantiers',   label: 'Chantiers',       href: 'chantiers.html',        icon: 'building' },
  { id: 'achats',      label: 'Achats & Devis',  href: 'achats.html',           icon: 'cart' },
  { id: 'maintenance', label: 'Maintenance',     href: 'maintenance.html',      icon: 'wrench' },
  { id: 'documents',   label: 'Documents',       href: 'documents.html',        icon: 'folder' },
  { id: 'support-it',  label: 'Support IT',      href: 'support-it.html',       icon: 'laptop' },
];

const SHORTCUT_ITEMS = [
  { label: 'Congés / RTT',     href: 'conges.html',          icon: 'calendar' },
  { label: 'Notes de frais',   href: 'notes-de-frais.html',  icon: 'euro' },
  { label: 'Pointage',         href: 'pointage.html',        icon: 'clock' },
  { label: 'Ticket IT',        href: 'ticket-it.html',       icon: 'ticket' },
];

const CURRENT_USER = {
  name: 'Alexandre Huard',
  role: 'Responsable Achat',
  initials: 'AH',
};

// ---------- Render helpers ----------
function renderShell() {
  // Topbar
  const headerSlot = document.getElementById('app-header');
  if (headerSlot) {
    headerSlot.outerHTML = `
      <header class="topbar">
        <button class="topbar-menu-btn" id="menuBtn" aria-label="Menu">${ICON.menu}</button>
        <a href="intranet.html" class="topbar-brand">
          <div class="brand-mark">${HUARD_LOGOMARK}</div>
          <div class="brand-text">
            <div class="brand-name">HUARD</div>
            <div class="brand-sub">GROUPE</div>
          </div>
        </a>
        <div class="topbar-search">
          ${ICON.search}
          <input type="search" placeholder="Rechercher un contact, un document, un chantier…">
        </div>
        <div class="topbar-actions">
          <button class="icon-btn" aria-label="Notifications" title="Notifications">${ICON.bell}<span class="dot"></span></button>
          <button class="icon-btn" aria-label="Paramètres" title="Paramètres">${ICON.settings}</button>
          <a href="index.html" class="icon-btn" aria-label="Déconnexion" title="Déconnexion">${ICON.logout}</a>
          <div class="topbar-user" id="userMenu">
            <div class="topbar-user-avatar">${CURRENT_USER.initials}</div>
            <div class="topbar-user-info">
              <div class="topbar-user-name">${CURRENT_USER.name}</div>
              <div class="topbar-user-role">${CURRENT_USER.role}</div>
            </div>
          </div>
        </div>
      </header>
    `;
  }

  // Sidebar (uses #app-nav slot)
  const navSlot = document.getElementById('app-nav');
  if (navSlot) {
    const active = document.body.dataset.page || '';
    const mainNav = NAV_ITEMS.map(i => `
      <a href="${i.href}" class="sidebar-link ${i.id === active ? 'active' : ''}">
        ${ICON[i.icon] || ''}
        <span>${i.label}</span>
      </a>
    `).join('');
    const shortcuts = SHORTCUT_ITEMS.map(i => `
      <a href="${i.href}" class="sidebar-link">
        ${ICON[i.icon] || ''}
        <span>${i.label}</span>
      </a>
    `).join('');
    navSlot.outerHTML = `
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-section">Navigation</div>
        ${mainNav}
        <div class="sidebar-section">Mes raccourcis</div>
        ${shortcuts}
        <div class="sidebar-footer">
          <div class="label">Astreinte 24/7</div>
          <div class="title">Alarme · poste 6129</div>
          <div class="num">RTM · poste 6128</div>
        </div>
      </aside>
      <div class="sidebar-backdrop" id="sidebarBackdrop"></div>
    `;
  }

  // Footer
  const footerSlot = document.getElementById('app-footer');
  if (footerSlot) {
    footerSlot.outerHTML = `
      <footer class="app-footer">
        © Groupe HUARD — Intranet interne · Accès réservé aux collaborateurs<br>
        Support : <a href="mailto:edurand@huard.fr">edurand@huard.fr</a> · poste 6113 (E. DURAND, Resp. Informatique)
      </footer>
    `;
  }

  // Mobile drawer wiring
  const menuBtn = document.getElementById('menuBtn');
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('sidebarBackdrop');
  if (menuBtn && sidebar && backdrop) {
    const toggle = () => {
      sidebar.classList.toggle('open');
      backdrop.classList.toggle('open');
    };
    menuBtn.addEventListener('click', toggle);
    backdrop.addEventListener('click', toggle);
  }

  // Inject SVG icons into any element with [data-icon]
  document.querySelectorAll('[data-icon]').forEach(el => {
    const name = el.dataset.icon;
    if (ICON[name]) el.innerHTML = ICON[name];
  });
}

// ---------- Staff directory (single source of truth) ----------
const HUARD_STAFF = [
  // Administratif / Direction
  { ext: 6100, tel: '01 73 23 61 00', nom: 'Standard',            role: 'Accueil téléphonique',         service: 'Administratif' },
  { ext: 6101, tel: '01 73 23 61 01', nom: 'Frédéric HUARD',      role: 'Direction',                    service: 'Administratif' },
  { ext: 6132, tel: '01 73 23 61 32', nom: 'Ricardo BIZARRO',     role: '',                              service: 'Administratif' },
  { ext: 6117, tel: '01 73 23 61 17', nom: 'Alain POTDEVIN',      role: '',                              service: 'Administratif' },
  { ext: 6103, tel: '01 73 23 61 03', nom: 'Sylvie BRICAUD',      role: 'Responsable Comptabilité',     service: 'Comptabilité' },
  { ext: 6106, tel: '01 73 23 61 06', nom: 'Liliane NGO',         role: 'Comptabilité',                 service: 'Comptabilité' },
  { ext: 6125, tel: '01 73 23 61 25', nom: 'Xavier HACOT',        role: 'Comptabilité',                 service: 'Comptabilité' },
  { ext: 6104, tel: '01 73 23 61 04', nom: 'Thibaut PLOUCHART',   role: 'Achat',                         service: 'Achats' },
  { ext: 6105, tel: '01 73 23 61 05', nom: 'Alexandre HUARD',     role: 'Responsable Achat',            service: 'Achats' },
  { ext: 6108, tel: '01 73 23 61 08', nom: 'Pascal JAPY',         role: '',                              service: 'Administratif' },
  { ext: 6109, tel: '01 73 23 61 09', nom: 'Salle de réunion',    role: '',                              service: 'Administratif' },
  // Sûreté / CFA
  { ext: 6120, tel: '01 73 23 61 20', nom: 'Valérie GUYOT',       role: '',                              service: 'Sûreté / CFA' },
  { ext: 6122, tel: '01 73 23 61 22', nom: 'Maxime HUGUET',       role: 'Responsable Sûreté / CFA',     service: 'Sûreté / CFA' },
  { ext: 6124, tel: '01 73 23 61 24', nom: 'Sixto SAAVEDRA',      role: 'Responsable Technique',         service: 'Sûreté / CFA' },
  { ext: 6129, tel: '01 73 23 61 29', nom: 'Astreinte Alarme',    role: 'Astreinte 24/7',                service: 'Sûreté / CFA' },
  // Courant fort / CFO
  { ext: 6126, tel: '01 73 23 61 26', nom: 'Philippe SALLE',      role: 'Responsable Courant Fort',     service: 'Courant Fort / CFO' },
  { ext: 6121, tel: '01 73 23 61 21', nom: 'Philippe LEVEQUE',    role: 'Conducteur de travaux',         service: 'Courant Fort / CFO' },
  { ext: 6123, tel: '01 73 23 61 23', nom: 'Fernando FIDALGO',    role: 'Conducteur de travaux',         service: 'Courant Fort / CFO' },
  { ext: 6115, tel: '01 73 23 61 15', nom: 'Carla VAZ',           role: 'Assistante CFO / CVC',         service: 'Courant Fort / CFO' },
  { ext: 6133, tel: '01 73 23 61 33', nom: 'Ernson BEAUSEJOUR',   role: 'Bureau d\'études',              service: 'Courant Fort / CFO' },
  // Informatique
  { ext: 6113, tel: '01 73 23 61 13', nom: 'Eric DURAND',         role: 'Responsable Informatique',      service: 'Informatique' },
  { ext: 6111, tel: '01 73 23 61 11', nom: 'Laetitia GEORGE',     role: 'Assistante Informatique / RTM', service: 'Informatique' },
  // Génie Climatique / CVC
  { ext: 6131, tel: '01 73 23 61 31', nom: 'Eric ANACLET',        role: 'Responsable CVC',               service: 'Génie Climatique / CVC' },
  { ext: 6130, tel: '01 73 23 61 30', nom: 'Stéphanie RIPOCHE',   role: 'Assistante CVC',                service: 'Génie Climatique / CVC' },
  { ext: 6107, tel: '01 73 23 61 07', nom: 'M. CELANO',           role: 'Conducteur de travaux CVC',    service: 'Génie Climatique / CVC' },
  // RTM Neophone
  { ext: 6118, tel: '01 73 23 61 18', nom: 'Jacky HENG',          role: 'Responsable RTM Neophone',     service: 'RTM Neophone' },
  { ext: 6116, tel: '01 73 23 61 16', nom: 'David GEIGER',        role: 'Support RTM',                   service: 'RTM Neophone' },
  { ext: 6128, tel: '01 73 23 61 28', nom: 'Astreinte RTM',       role: 'Astreinte 24/7',                service: 'RTM Neophone' },
  // Fax
  { ext: 6139, tel: '01 73 23 61 39', nom: 'Fax HUARD',           role: 'Fax',                           service: 'Administratif' },
  { ext: 6137, tel: '01 73 23 61 37', nom: 'Fax RTM Neophone',    role: 'Fax',                           service: 'RTM Neophone' },
];

function huardManagers() {
  return HUARD_STAFF.filter(p => /Responsable/i.test(p.role));
}
function initialsOf(nom) {
  return nom.split(/\s+/).map(x => x[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
}
function telHref(tel) {
  return 'tel:+33' + tel.replace(/[^0-9]/g, '').slice(1);
}

// ---------- Boot ----------
document.addEventListener('DOMContentLoaded', renderShell);
