// voice-input.js — голосовой ввод результатов (Web Speech API)
//
// Одна кнопка-переключатель «🎤 Голосовой ввод» на странице. Когда она
// включена, наведение курсора на любое поле из списка запускает
// распознавание речи (русский язык) для этого поля: распознанный текст
// подставляется в поле и генерируется событие 'input' — дальнейшее
// форматирование (цифры → ЧЧ:ММ:СС.сс или М:СС.сс) делает уже
// существующий на странице обработчик.
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

    var inputs = Array.prototype.slice.call(document.querySelectorAll(selector))
      .filter(function (inp) { return !inp.disabled; });
    if (!inputs.length) return;

    var активен = false;
    var rec = null;
    var текущееПоле = null;

    var toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'voice-input-toggle no-print';
    toggle.textContent = '🎤 Голосовой ввод';
    document.body.appendChild(toggle);

    function остановить() {
      if (rec) {
        var r = rec;
        rec = null;
        r.onresult = null;
        r.onerror = null;
        r.onend = null;
        r.abort();
      }
      if (текущееПоле) {
        текущееПоле.classList.remove('voice-input-active');
        текущееПоле = null;
      }
    }

    function начатьДляПоля(inp) {
      остановить();
      текущееПоле = inp;
      inp.classList.add('voice-input-active');

      rec = new SpeechRecognitionCtor();
      rec.lang = 'ru-RU';
      rec.interimResults = false;
      rec.maxAlternatives = 1;

      rec.onresult = function (ev) {
        var text = ev.results[0][0].transcript || '';
        inp.value = text;
        inp.dispatchEvent(new Event('input', { bubbles: true }));
      };
      rec.onerror = function () {
        остановить();
      };
      rec.onend = function () {
        if (текущееПоле === inp) {
          inp.classList.remove('voice-input-active');
          текущееПоле = null;
        }
        rec = null;
      };
      rec.start();
    }

    toggle.addEventListener('click', function () {
      активен = !активен;
      toggle.classList.toggle('on', активен);
      if (!активен) остановить();
    });

    inputs.forEach(function (inp) {
      inp.addEventListener('mouseenter', function () {
        if (!активен || текущееПоле === inp) return;
        начатьДляПоля(inp);
      });
      inp.addEventListener('mouseleave', function () {
        if (текущееПоле === inp) остановить();
      });
    });
  };
})();
