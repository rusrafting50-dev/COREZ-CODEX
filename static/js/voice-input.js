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
      rec.onerror = function (ev) {
        console.warn('Голосовой ввод: ошибка распознавания —', ev.error);
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

    function включить() {
      активен = true;
      toggle.classList.add('on');
    }

    toggle.addEventListener('click', function () {
      if (активен) {
        активен = false;
        toggle.classList.remove('on');
        остановить();
        return;
      }
      // Разрешение на микрофон запрашиваем по прямому клику — иначе браузер
      // может молча отклонить запрос, сделанный позже по наведению курсора.
      if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
        navigator.mediaDevices.getUserMedia({ audio: true }).then(function (stream) {
          stream.getTracks().forEach(function (t) { t.stop(); });
          включить();
        }).catch(function (err) {
          console.warn('Голосовой ввод: доступ к микрофону не получен —', err);
          alert('Не удалось получить доступ к микрофону.\nПроверьте значок разрешений рядом с адресом сайта в браузере и дайте доступ к микрофону для этой страницы.');
        });
      } else {
        включить();
      }
    });

    inputs.forEach(function (inp) {
      inp.addEventListener('mouseenter', function () {
        if (!активен || текущееПоле === inp) return;
        начатьДляПоля(inp);
      });
    });
  };
})();
