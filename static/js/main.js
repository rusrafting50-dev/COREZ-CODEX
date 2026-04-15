// main.js — вспомогательные скрипты СЧИТАЛКА

// Подтверждение удаления (используется через onsubmit="return confirm(...)")
// — уже встроено в шаблоны через атрибуты

// Подсветка строк таблицы по месту
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('tr[data-place]').forEach(function (row) {
    var place = parseInt(row.dataset.place, 10);
    if (place === 1) row.classList.add('place-1');
    else if (place === 2) row.classList.add('place-2');
    else if (place === 3) row.classList.add('place-3');
  });
});
