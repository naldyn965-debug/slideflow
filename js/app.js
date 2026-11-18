import { $, h, icon } from './lib/dom.js';
import { createStore } from './store/index.js';
import { toast } from './ui/toast.js';
import { errorState } from './ui/states.js';
import { dashboard } from './views/dashboard.js';
import { projects } from './views/projects.js';
import { project } from './views/project.js';
import { templates } from './views/templates.js';
import { signInView } from './views/signin.js';

const NAV = [['#/', 'Dashboard', 'dashboard'], ['#/projects', 'Projects', 'folder'], ['#/templates', 'Templates', 'slides']];
const ROUTES = [
  [/^#?\/?$/, 'Dashboard', (c, r) => dashboard(c, r)],
  [/^#\/projects$/, 'Projects', (c, r) => projects(c, r)],
  [/^#\/projects\/([\w-]+)$/, 'Project', (c, r, m) => project(c, r, m[1])],
  [/^#\/templates$/, 'Templates', (c, r) => templates(c, r)],
];

async function boot() {
  const app = $('app');
  app.replaceChildren(h('p', { class: 'boot' }, 'Loading SlideFlow…'));
  let store;
  try { store = await createStore(); }
  catch (e) { console.error(e); return app.replaceChildren(h('div', { class: 'boot' }, errorState(e, () => location.reload()))); }

  const ctx = { store, toast };
  let main = null;
  let navLinks = [];

  const route = () => {
    if (!main) return;
    const hash = location.hash || '#/';
    const hit = ROUTES.map(([re, title, fn]) => [re.exec(hash), title, fn]).find(([m]) => m) || [null, 'Dashboard', ROUTES[0][2]];
    const [match, title, fn] = hit;
    document.title = `${title} · SlideFlow`;
    const section = hash.replace(/^#\/?/, '').split('/')[0];
    navLinks.forEach(([a, href]) => (href.replace(/^#\/?/, '') === section ? a.setAttribute('aria-current', 'page') : a.removeAttribute('aria-current')));
    const view = h('div', { class: 'view' });
    main.replaceChildren(view);
    fn(ctx, view, match || []);
    main.focus({ preventScroll: true });
    window.scrollTo(0, 0);
  };

  const mount = (user) => {
    ctx.user = user;
    if (!user) { main = null; return app.replaceChildren(signInView(ctx)); }
    navLinks = NAV.map(([href, label, ic]) => [h('a', { href }, icon(ic), label), href]);
    main = h('main', { id: 'main', tabindex: '-1' });
    app.replaceChildren(h('div', { class: 'shell' },
      h('aside', { class: 'side' },
        h('div', { class: 'brand' }, h('span', { class: 'mark' }), 'SlideFlow'),
        h('nav', { 'aria-label': 'Main' }, navLinks.map(([a]) => a)),
        h('div', { class: 'side-foot' },
          store.mode === 'local' && h('p', { class: 'mode' }, 'Local mode. Data stays in this browser until Firebase is configured.'),
          store.mode === 'firebase' && h('div', { class: 'user' }, h('span', {}, user.name), h('button', { class: 'btn btn-quiet', type: 'button', onclick: () => store.session.signOut() }, 'Sign out')))),
      main));
    route();
  };

  window.addEventListener('hashchange', route);
  store.session.onChange(mount);
}

boot();
