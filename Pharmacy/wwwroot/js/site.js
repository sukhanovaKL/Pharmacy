(function (window, document) {
    'use strict';

    const PharmacyApp = {
        init() {
            this.bindConfirmDeletes();
            this.bindFlashAutoHide();
            this.bindClientTableFilter();
            this.bindFormValidationHints();
        },

        bindConfirmDeletes() {
            document.querySelectorAll('form.js-confirm-delete').forEach((form) => {
                form.addEventListener('submit', (event) => {
                    const message = form.getAttribute('data-confirm') || 'Подтвердите удаление';
                    if (!window.confirm(message)) {
                        event.preventDefault();
                    }
                });
            });
        },

        bindFlashAutoHide() {
            document.querySelectorAll('.js-flash').forEach((el) => {
                setTimeout(() => {
                    el.classList.add('fade');
                    setTimeout(() => el.remove(), 400);
                }, 4000);
            });
        },

        bindClientTableFilter() {
            const searchBox = document.getElementById('searchBox');
            const table = document.getElementById('dataTable');
            if (!searchBox || !table || !searchBox.closest('form.js-live-filter')) {
                return;
            }

            searchBox.addEventListener('input', () => {
                const q = searchBox.value.trim().toLowerCase();
                table.querySelectorAll('tbody tr').forEach((row) => {
                    const text = row.textContent.toLowerCase();
                    row.style.display = !q || text.includes(q) ? '' : 'none';
                });
            });
        },

        bindFormValidationHints() {
            document.querySelectorAll('form.js-validate-form').forEach((form) => {
                form.addEventListener('submit', (event) => {
                    if (!form.checkValidity()) {
                        event.preventDefault();
                        event.stopPropagation();
                        const firstInvalid = form.querySelector(':invalid');
                        if (firstInvalid) {
                            firstInvalid.focus();
                        }
                    }
                    form.classList.add('was-validated');
                });
            });
        },

        initSaleForm() {
            const container = document.getElementById('saleLines');
            const template = document.getElementById('saleLineTemplate');
            const addBtn = document.getElementById('addSaleLine');
            if (!container || !template || !addBtn) {
                return;
            }

            const reindex = () => {
                container.querySelectorAll('.sale-line').forEach((row, index) => {
                    row.querySelectorAll('select, input').forEach((input) => {
                        const name = input.getAttribute('name');
                        if (!name) return;
                        input.setAttribute('name', name.replace(/Lines\[\d+\]/, `Lines[${index}]`));
                    });
                });
            };

            addBtn.addEventListener('click', () => {
                const index = container.querySelectorAll('.sale-line').length;
                const html = template.innerHTML.replaceAll('__index__', String(index));
                container.insertAdjacentHTML('beforeend', html);
            });

            container.addEventListener('click', (event) => {
                const btn = event.target.closest('.js-remove-line');
                if (!btn) return;
                const row = btn.closest('.sale-line');
                if (!row) return;
                if (container.querySelectorAll('.sale-line').length <= 1) {
                    window.alert('Должна остаться хотя бы одна строка.');
                    return;
                }
                row.remove();
                reindex();
            });
        }
    };

    document.addEventListener('DOMContentLoaded', () => PharmacyApp.init());
    window.PharmacyApp = PharmacyApp;
})(window, document);
