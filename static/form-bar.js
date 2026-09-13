/*
 * The sticky action bar at the foot of a document form.
 *
 * It exists so the primary action is reachable without scrolling to the
 * bottom of a long form, and so the two things you would scroll back up
 * to check - who this is for, and whether anything is still wrong - are
 * visible while you work.
 *
 * Read-only: it reports what the form already says. All validation stays
 * where it was, in the browser's own constraint checking and then on the
 * server; nothing here decides whether a form may be submitted.
 */
(function () {
  var form = document.getElementById('handover-form');
  var bar = document.getElementById('form-bar');
  if (!form || !bar) return;

  var summary = document.getElementById('bar-summary');
  var status = document.getElementById('bar-status');
  var button = document.getElementById('submit-btn');
  var base = summary ? summary.textContent.trim() : '';

  function value(name) {
    var el = form.querySelector('[name="' + name + '"]');
    return el ? el.value.trim() : '';
  }

  /* The model of whichever device section comes first - enough to tell
     two half-filled forms apart without naming every field. */
  function device() {
    var el = form.querySelector('input[name$="model"], input[name$="brand"]');
    return el ? el.value.trim() : '';
  }

  /* The format rule lives in the input's title. It is shown under the
     field only while that field is actually wrong, rather than sitting
     grey under every field the way the old form did. */
  function setFieldError(el, message) {
    var holder = el.closest('.field');
    if (!holder) return;
    var note = holder.querySelector('.field-error');
    if (!message) {
      if (note) note.remove();
      return;
    }
    if (!note) {
      note = document.createElement('p');
      note.className = 'field-error';
      note.innerHTML = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none"' +
        ' stroke="currentColor" stroke-width="2.2" stroke-linecap="round">' +
        '<circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg><span></span>';
      holder.appendChild(note);
    }
    note.querySelector('span').textContent = message;
  }

  function refresh() {
    var who = value('name');
    var what = device();
    summary.textContent = [who, what].filter(Boolean).join('  ·  ') || base;

    /* Only count fields the person has actually engaged with, so an
       untouched form is never accused of having problems. */
    var bad = 0;
    form.querySelectorAll('input[name]').forEach(function (el) {
      var wrong = el.value.trim() && el.willValidate && !el.checkValidity();
      if (wrong) bad++;
      setFieldError(el, wrong ? (el.title || 'That is not the right format.') : '');
    });
    if (bad) {
      status.textContent = bad === 1 ? '1 field needs attention'
                                     : bad + ' fields need attention';
      status.className = 'bar-warn';
    } else {
      status.textContent = '';
      status.className = '';
    }
  }

  form.addEventListener('input', refresh);
  form.addEventListener('change', refresh);
  form.addEventListener('submit', function () {
    button.disabled = true;
    button.classList.add('loading');
    var label = button.querySelector('.btn-label');
    if (label) label.textContent = label.textContent.indexOf('Save') === 0 ? 'Saving…' : 'Generating…';
  });
  refresh();
})();
