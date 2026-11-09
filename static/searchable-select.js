/*
 * Selector con búsqueda (sin dependencias).
 *
 * Uso: agregar el atributo data-searchable a un <select>.
 *   <select name="product_id" data-searchable data-placeholder="Buscar..."> ... </select>
 *
 * El <select> original se mantiene (oculto) y es el que se envía con el
 * formulario, así que el backend no cambia. Las opciones se ordenan
 * alfabéticamente y el texto escrito filtra por cualquier parte del texto
 * de la opción (código, nombre, descripción), sin importar mayúsculas ni
 * tildes. Varias palabras se buscan todas a la vez.
 */
(function () {
    'use strict';

    function normalize(text) {
        return (text || '')
            .toString()
            .normalize('NFD')
            .replace(/[̀-ͯ]/g, '')
            .toLowerCase()
            .trim();
    }

    function enhance(select) {
        if (select.dataset.searchableReady) return;
        select.dataset.searchableReady = '1';

        var collator = new Intl.Collator('es', { numeric: true, sensitivity: 'base' });

        var options = Array.prototype.slice.call(select.options)
            .filter(function (o) { return o.value !== ''; })
            .map(function (o) {
                return { value: o.value, label: o.text.trim(), key: normalize(o.text) };
            })
            .sort(function (a, b) { return collator.compare(a.label, b.label); });

        var wrapper = document.createElement('div');
        wrapper.className = 'ss-wrap';

        var input = document.createElement('input');
        input.type = 'text';
        input.className = 'ss-input';
        input.autocomplete = 'off';
        input.placeholder = select.dataset.placeholder || 'Escriba para buscar...';

        var list = document.createElement('div');
        list.className = 'ss-list';
        list.setAttribute('role', 'listbox');

        select.parentNode.insertBefore(wrapper, select);
        wrapper.appendChild(input);
        wrapper.appendChild(list);
        wrapper.appendChild(select);
        select.classList.add('ss-native');
        select.tabIndex = -1;

        var MAX_RESULTS = 200;
        var visible = [];
        var active = -1;

        function selectedLabel() {
            var opt = select.options[select.selectedIndex];
            return opt && opt.value !== '' ? opt.text.trim() : '';
        }

        function render(query) {
            var terms = normalize(query).split(/\s+/).filter(Boolean);
            visible = options.filter(function (o) {
                return terms.every(function (t) { return o.key.indexOf(t) !== -1; });
            });

            list.innerHTML = '';
            if (!visible.length) {
                var empty = document.createElement('div');
                empty.className = 'ss-empty';
                empty.textContent = 'Sin resultados';
                list.appendChild(empty);
            }
            visible.slice(0, MAX_RESULTS).forEach(function (o, i) {
                var item = document.createElement('div');
                item.className = 'ss-item';
                item.setAttribute('role', 'option');
                item.dataset.index = i;
                item.textContent = o.label;
                if (o.value === select.value) item.classList.add('ss-selected');
                list.appendChild(item);
            });
            if (visible.length > MAX_RESULTS) {
                var more = document.createElement('div');
                more.className = 'ss-empty';
                more.textContent = (visible.length - MAX_RESULTS) + ' más... siga escribiendo para filtrar';
                list.appendChild(more);
            }
            setActive(visible.length ? 0 : -1);
        }

        function setActive(i) {
            var items = list.querySelectorAll('.ss-item');
            if (active >= 0 && items[active]) items[active].classList.remove('ss-active');
            active = i;
            if (active >= 0 && items[active]) {
                items[active].classList.add('ss-active');
                items[active].scrollIntoView({ block: 'nearest' });
            }
        }

        function open() {
            wrapper.classList.add('ss-open');
            render(input.value === selectedLabel() ? '' : input.value);
        }

        function close() {
            wrapper.classList.remove('ss-open');
            input.value = selectedLabel();
        }

        function choose(i) {
            var o = visible[i];
            if (!o) return;
            select.value = o.value;
            select.dispatchEvent(new Event('change', { bubbles: true }));
            input.setCustomValidity('');
            close();
        }

        input.addEventListener('focus', function () { input.select(); open(); });
        input.addEventListener('input', function () {
            wrapper.classList.add('ss-open');
            render(input.value);
        });
        input.addEventListener('keydown', function (e) {
            var count = list.querySelectorAll('.ss-item').length;
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                if (!wrapper.classList.contains('ss-open')) open();
                else if (count) setActive(Math.min(active + 1, count - 1));
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                if (count) setActive(Math.max(active - 1, 0));
            } else if (e.key === 'Enter') {
                if (wrapper.classList.contains('ss-open')) {
                    e.preventDefault();
                    choose(active);
                }
            } else if (e.key === 'Escape') {
                close();
            } else if (e.key === 'Tab') {
                close();
            }
        });

        list.addEventListener('mousedown', function (e) {
            // Evita que el input pierda el foco antes de elegir.
            e.preventDefault();
            var item = e.target.closest('.ss-item');
            if (item) choose(Number(item.dataset.index));
        });

        document.addEventListener('mousedown', function (e) {
            if (!wrapper.contains(e.target)) close();
        });

        // Si el formulario se envía sin elegir, mostrar el aviso en el buscador.
        select.addEventListener('invalid', function (e) {
            e.preventDefault();
            input.setCustomValidity('Seleccione un producto de la lista');
            input.reportValidity();
        });
        input.addEventListener('input', function () { input.setCustomValidity(''); });

        input.value = selectedLabel();
    }

    function init(root) {
        (root || document).querySelectorAll('select[data-searchable]').forEach(enhance);
    }

    window.SearchableSelect = { init: init };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () { init(); });
    } else {
        init();
    }
})();
