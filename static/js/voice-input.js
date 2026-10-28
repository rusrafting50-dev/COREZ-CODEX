// voice-input.js — голосовой ввод результатов (Web Speech API)
//
// Добавляет кнопку 🎤 внутрь текстовых полей времени. По нажатию запускает
// распознавание речи (русский язык) и подставляет распознанный текст в поле,
// затем генерирует событие 'input' — дальнейшее форматирование (цифры →
// ЧЧ:ММ:СС.сс или М:СС.сс) делает уже существующий на странице обработчик.
//
// Работает только в Chrome/Edge (Web Speech API), требует подключения
// к интернету (распознавание идёт через сервис Google) и разрешения
// на доступ к микрофону.

(function () {
  function поддерживается() {
    return 'webkitSpeechRecognition' in window || 'SpeechRecognition' in window;
  }

  window.включитьГолосовойВвод = function (selector) {
    if (!поддерживается()) return;
    var SpeechRecognitionCtor = window.SpeechRecognition || window.webkitSpeechRecognition;

    document.querySelectorAll(selector).forEach(function (inp) {
      if (inp.dataset.voiceReady || inp.disabled) return;
      inp.dataset.voiceReady = '1';

      var wrap = document.createElement('span');
      wrap.className = 'voice-input-wrap';
      inp.parentNode.insertBefore(wrap, inp);
      wrap.appendChild(inp);

      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'voice-input-btn';
      btn.title = 'Голосовой ввод (русский)';
      btn.textContent = '🎤';
      wrap.appendChild(btn);

      var rec = null;

      btn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (rec) {
          rec.stop();
          return;
        }
        rec = new SpeechRecognitionCtor();
        rec.lang = 'ru-RU';
        rec.interimResults = false;
        rec.maxAlternatives = 1;
        btn.classList.add('listening');

        rec.onresult = function (ev) {
          var text = ev.results[0][0].transcript || '';
          inp.value = text;
          inp.dispatchEvent(new Event('input', { bubbles: true }));
          inp.focus();
        };
        rec.onerror = function () {
          btn.classList.remove('listening');
        };
        rec.onend = function () {
          btn.classList.remove('listening');
          rec = null;
        };
        rec.start();
      });
    });
  };
})();
