import { h } from '../lib/dom.js';
import { friendly } from '../ui/states.js';

export function signInView(ctx) {
  const msg = h('p', { class: 'form-error', role: 'alert' });
  const go = (kind) => async () => {
    msg.textContent = '';
    try { await ctx.store.session.signIn(kind); }
    catch (e) {
      if (['auth/popup-closed-by-user', 'auth/cancelled-popup-request'].includes(e.code)) return;
      console.error(e);
      msg.textContent = e.code === 'auth/operation-not-allowed' ? 'This sign-in method is not enabled in the Firebase console.' : friendly(e);
    }
  };
  return h('main', { class: 'signin' },
    h('div', { class: 'signin-card' },
      h('div', { class: 'brand' }, h('span', { class: 'mark' }), 'SlideFlow'),
      h('h1', {}, 'Sign in to your workspace'),
      h('p', { class: 'muted' }, 'Your templates and projects are private to your account.'),
      h('button', { class: 'btn btn-primary wide', type: 'button', onclick: go('google') }, 'Continue with Google'),
      h('button', { class: 'btn wide', type: 'button', onclick: go('guest') }, 'Continue as guest'),
      h('p', { class: 'muted fine' }, 'Guest data stays with this browser. Clearing site data removes access to it.'),
      msg));
}
