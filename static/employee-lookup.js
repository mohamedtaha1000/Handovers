/*
 * "Reuse a previous employee" suggestions.
 *
 * Attached both to a labelled search box at the top of the employee
 * section and to the full-name and employee-code fields themselves, so
 * it is findable without being in the way: the box says what it does,
 * and typing a name you have used before in the ordinary field offers it
 * back anyway. Typing a new name behaves exactly as it always did.
 *
 * Only the employee half of the form is ever filled in. Equipment
 * details are deliberately left alone - the serial number of the laptop
 * someone was issued last year must not follow them onto a new document.
 */
(function () {
  var EMPLOYEE_KEYS = ['name', 'department', 'role', 'mobile', 'email', 'code', 'govid'];
  var inputs = document.querySelectorAll('[data-employee-lookup]');
  if (!inputs.length) return;

  var cache = null;
  var box = null;
  var active = null;

  function load() {
    if (cache) return Promise.resolve(cache);
    return fetch('/api/employees', { credentials: 'same-origin' })
      .then(function (r) { return r.ok ? r.json() : { employees: [] }; })
      .then(function (d) { cache = d.employees || []; return cache; })
      .catch(function () { return []; });   // offline or logged out: just behave like a plain box
  }

  function close() {
    if (box) { box.remove(); box = null; }
    active = null;
  }

  /* The form shows the part before "@stm.com.eg"; the stored value is
     the full address, so trim it back or it would round-trip wrongly. */
  function emailLocalPart(value) {
    return (value || '').split('@')[0];
  }

  function apply(person) {
    EMPLOYEE_KEYS.forEach(function (key) {
      var field = document.querySelector('[name="' + key + '"]');
      if (!field) return;
      var value = person[key] || '';
      field.value = key === 'email' ? emailLocalPart(value) : value;
      /* Let any validation styling re-evaluate against the new value. */
      field.dispatchEvent(new Event('input', { bubbles: true }));
    });
    var search = document.getElementById('employee-search');
    if (search) search.value = person.name || '';
    close();
    /* Straight on to the equipment, which is the only part left to fill. */
    var form = document.getElementById('handover-form');
    var firstDevice = form && form.querySelector('.card.section + .card.section input');
    if (firstDevice) firstDevice.focus();
    if (form) form.dispatchEvent(new Event('input', { bubbles: true }));
  }

  function render(input, matches) {
    close();
    if (!matches.length) return;
    box = document.createElement('div');
    box.className = 'lookup-menu';
    box.setAttribute('role', 'listbox');
    matches.forEach(function (person, i) {
      var item = document.createElement('button');
      item.type = 'button';
      item.className = 'lookup-item';
      item.setAttribute('role', 'option');
      item.dataset.index = i;
      var bits = [person.code, person.department].filter(Boolean).join(' · ');
      item.innerHTML = '<span class="lookup-name"></span>' +
                       (bits ? '<span class="lookup-meta"></span>' : '');
      item.querySelector('.lookup-name').textContent = person.name;
      if (bits) item.querySelector('.lookup-meta').textContent = bits;
      item.addEventListener('mousedown', function (e) { e.preventDefault(); apply(person); });
      box.appendChild(item);
    });
    (input.closest('.field, .lookup-field') || input.parentNode).appendChild(box);
    active = input;
  }

  function suggest(input) {
    var q = input.value.trim().toLowerCase();
    if (q.length < 2) { close(); return; }
    load().then(function (people) {
      if (document.activeElement !== input) return;
      var matches = people.filter(function (p) {
        return (p.name || '').toLowerCase().indexOf(q) !== -1 ||
               (p.code || '').toLowerCase().indexOf(q) !== -1;
      }).slice(0, 6);
      /* Nothing to offer if the only match is exactly what is typed. */
      if (matches.length === 1 &&
          (matches[0].name || '').toLowerCase() === q) { close(); return; }
      render(input, matches);
    });
  }

  inputs.forEach(function (input) {
    input.addEventListener('input', function () { suggest(input); });
    input.addEventListener('focus', function () { suggest(input); });
    input.addEventListener('blur', function () { setTimeout(close, 120); });
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') { close(); return; }
      if (!box || (e.key !== 'ArrowDown' && e.key !== 'ArrowUp' && e.key !== 'Enter')) return;
      var items = Array.prototype.slice.call(box.querySelectorAll('.lookup-item'));
      var current = items.indexOf(box.querySelector('.lookup-item.on'));
      if (e.key === 'Enter') {
        if (current > -1) { e.preventDefault(); items[current].dispatchEvent(new Event('mousedown')); }
        return;
      }
      e.preventDefault();
      var next = e.key === 'ArrowDown' ? current + 1 : current - 1;
      if (next < 0) next = items.length - 1;
      if (next >= items.length) next = 0;
      items.forEach(function (el) { el.classList.remove('on'); });
      items[next].classList.add('on');
    });
  });

  document.addEventListener('click', function (e) {
    if (box && active && !box.contains(e.target) && e.target !== active) close();
  });
})();
