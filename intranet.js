// Shared header, nav and footer for the HUARD intranet.
// Each page sets <body data-page="..."> to highlight the active tab.

const HUARD_LOGO_SVG = `
<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <rect x="2" y="2" width="60" height="60" rx="10" fill="#1F2F4D"/>
  <path d="M16 14 H22 V28 H32 V14 H38 V50 H32 V34 H22 V50 H16 Z" fill="#FFFFFF"/>
  <rect x="42" y="14" width="6" height="36" fill="#5D81A6"/>
</svg>`;

const NAV_ITEMS = [
  { id: 'accueil',     label: 'Accueil',          href: 'intranet.html' },
  { id: 'annuaire',    label: 'Annuaire',         href: 'annuaire.html' },
  { id: 'documents',   label: 'Documents',        href: 'documents.html' },
  { id: 'rh',          label: 'RH',               href: 'rh.html' },
  { id: 'chantiers',   label: 'Chantiers',        href: 'chantiers.html' },
  { id: 'achats',      label: 'Achats & Devis',   href: 'achats.html' },
  { id: 'maintenance', label: 'Maintenance',      href: 'maintenance.html' },
  { id: 'support-it',  label: 'Support IT',       href: 'support-it.html' },
];

const CURRENT_USER = {
  name: 'Alexandre Huard',
  role: 'Responsable Achat',
  initials: 'AH',
};

function renderHeader() {
  const slot = document.getElementById('app-header');
  if (!slot) return;
  slot.innerHTML = `
    <header class="app">
      <a href="intranet.html" class="brand">
        <div class="brand-logo">${HUARD_LOGO_SVG}</div>
        <div>
          <div class="brand-title">HUARD GROUPE</div>
          <div class="brand-subtitle">Intranet interne</div>
        </div>
      </a>
      <div class="user-box">
        <div style="text-align:right">
          <div class="user-name">${CURRENT_USER.name}</div>
          <div class="user-role">${CURRENT_USER.role}</div>
        </div>
        <div class="user-avatar">${CURRENT_USER.initials}</div>
        <a href="index.html" class="logout">Déconnexion</a>
      </div>
    </header>
  `;
}

function renderNav() {
  const slot = document.getElementById('app-nav');
  if (!slot) return;
  const active = document.body.dataset.page || '';
  slot.innerHTML = `
    <nav class="main">
      ${NAV_ITEMS.map(i => `<a href="${i.href}" class="${i.id === active ? 'active' : ''}">${i.label}</a>`).join('')}
    </nav>
  `;
}

function renderFooter() {
  const slot = document.getElementById('app-footer');
  if (!slot) return;
  slot.innerHTML = `
    <footer class="app">
      © Groupe HUARD — Intranet interne · Accès réservé aux collaborateurs<br>
      Pour toute question : <a href="mailto:edurand@huard.fr">edurand@huard.fr</a> · poste 6113 (Support Informatique)
    </footer>
  `;
}

// Annuaire data — source unique pour toutes les pages
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

document.addEventListener('DOMContentLoaded', () => {
  renderHeader();
  renderNav();
  renderFooter();
});
