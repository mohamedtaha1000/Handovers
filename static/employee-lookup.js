/*
 * "Reuse someone already on file" suggestions.
 *
 * Attached both to a labelled search box at the top of the employee
 * section and to the full-name and employee-code fields themselves, so
 * it is findable without being in the way: the box says what it does,
 * and typing a name you have used before in the ordinary field offers it
 * back anyway. Typing a new name behaves exactly as it always did.
 *
 * On a handover form only the employee half is ever filled in. Equipment
 * details are deliberately left alone - the serial number of the laptop
 * someone was issued last year must not follow them onto a new document.
 *
 * The leaver page wants the opposite: its whole point is the equipment
 * that person is still holding. So which fields get filled, where the
 * suggestions come from, and whether the email is trimmed to its
 * username are all read off the input, and each page says what it needs:
 *
 *   data-employee-lookup              the hook (required)
 *   data-lookup-url    default /api/employees
 *   data-lookup-keys   comma-separated; default the employee half
 *   data-lookup-email  "full" to keep name@stm.com.eg, default trims it
 */
(function () {
  var DEFAULT_KEYS = ['name', 'department', 'role', 'mobile', 'email', 'code', 'govid'];
  var inputs = document.querySelectorAll('[data-employee-lookup]');
  if (!inputs.length) return;

  function settings(input) {
    var keys = (input.dataset.lookupKeys || '').split(',')
      .map(function (k) { return k.trim(); }).filter(Boolean);
    return {
      url: input.dataset.lookupUrl || '/api/employees',
      keys: keys.length ? keys : DEFAULT_KEYS,
      fullEmail: input.dataset.lookupEmail === 'full'
    };
  }

  var cache = {};
  var box = null;
  var active = null;

  function load(url) {
    if (cache[url]) return Promise.resolve(cache[url]);
    return fetch(url, { credentials: 'same-origin' })
      .then(function (r) { return r.ok ? r.json() : { employees: [] }; })
      .then(function (d) { cache[url] = d.employees || []; return cache[url]; })
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

  function apply(person, opts, source) {
    opts.keys.forEach(function (key) {
      var field = document.querySelector('[name="' + key + '"]');
      if (!field) return;
      var value = person[key] || '';
      field.value = (key === 'email' && !opts.fullEmail) ? emailLocalPart(value) : value;
      /* Let any validation styling re-evaluate against the new value. */
      field.dispatchEvent(new Event('input', { bubbles: true }));
    });
    /* A search box carries no name attribute, so it is not one of the
       fields just filled: show who was picked in it. */
    if (source && !source.getAttribute('name')) source.value = person.name || '';
    close();
    /* Straight on to the equipment, which is the only part left to fill. */
    var form = document.getElementById('handover-form');
    var firstDevice = form && form.querySelector('.card.section + .card.section input');
    if (firstDevice) firstDevice.focus();
    if (form) form.dispatchEvent(new Event('input', { bubbles: true }));
  }

  function render(input, matches, opts) {
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
      item.addEventListener('mousedown', function (e) { e.preventDefault(); apply(person, opts, input); });
      box.appendChild(item);
    });
    (input.closest('.field, .lookup-field') || input.parentNode).appendChild(box);
    active = input;
  }

  function suggest(input) {
    var opts = settings(input);
    var q = input.value.trim().toLowerCase();
    if (q.length < 2) { close(); return; }
    load(opts.url).then(function (people) {
      if (document.activeElement !== input) return;
      var matches = people.filter(function (p) {
        return (p.name || '').toLowerCase().indexOf(q) !== -1 ||
               (p.code || '').toLowerCase().indexOf(q) !== -1;
      }).slice(0, 6);
      /* Nothing to offer if the only match is exactly what is typed. */
      if (matches.length === 1 &&
          (matches[0].name || '').toLowerCase() === q) { close(); return; }
      render(input, matches, opts);
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
