    /* ================================================================
       COLE AQUI a URL do seu App da Web do Google (veja INSTRUCOES.md).
       Enquanto estiver com o texto abaixo, o formulário NÃO envia de
       verdade — ele só mostra um aviso de teste.
       ================================================================ */
    var URL_PLANILHA = "https://script.google.com/macros/s/AKfycbxzYmuJCHvxb48mXnoBojttFlAGKb5w9ENPeg1obbQvzulYtsldTp8FlAKXoLCbZoUZ_Q/exec";

    // ---------- abertura da carta ----------
    (function(){
      var envelope = document.getElementById('envelope');
      var overlay  = document.getElementById('cartaOverlay');
      var aberto = false;
      function abrirCarta(){
        if(aberto) return;
        aberto = true;
        envelope.classList.add('aberto');
        // a cartinha fica na tela um tempinho antes do convite aparecer
        setTimeout(function(){
          overlay.classList.add('sumir');
          document.body.classList.remove('carta-fechada');
        }, 2200);
        setTimeout(function(){ overlay.remove(); }, 3000);
      }
      envelope.addEventListener('click', abrirCarta);
      envelope.addEventListener('keydown', function(e){
        if(e.key === 'Enter' || e.key === ' ') abrirCarta();
      });
    })();

    (function(){
      var video = document.getElementById('convite-video');
      var form  = document.getElementById('form');
      var lock  = document.getElementById('locknote');
      var bar   = document.getElementById('progressbar');
      var btn   = document.getElementById('btnSubmit');
      var msg   = document.getElementById('msg');
      var unlocked = false;

      function unlock(){
        if(unlocked) return;
        unlocked = true;
        form.classList.remove('locked');
        form.classList.add('reveal');
        lock.style.display = 'none';
        document.getElementById('eventoDetalhes').style.display = 'block';
        form.scrollIntoView({behavior:'smooth', block:'center'});
      }

      // barra de progresso conforme assiste
      video.addEventListener('timeupdate', function(){
        if(video.duration){
          bar.style.width = Math.min(100,(video.currentTime/video.duration)*100) + '%';
        }
      });
      // libera o formulário quando o vídeo termina
      video.addEventListener('ended', unlock);

      // ---------- envio da inscrição ----------
      var nome  = document.getElementById('nome');
      var wpp   = document.getElementById('wpp');
      var email = document.getElementById('email');

      function emailValido(v){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v); }

      // mascara automatica do WhatsApp: (00) 00000-0000
      wpp.addEventListener('input', function(){
        var d = wpp.value.replace(/\D/g,'').slice(0,11);
        var f = '';
        if(d.length > 0) f = '(' + d.slice(0,2);
        if(d.length >= 3) f += ') ' + d.slice(2,7);
        if(d.length >= 8) f += '-' + d.slice(7);
        wpp.value = f;
      });
      function limpaErros(){ [nome,wpp,email].forEach(function(el){ el.classList.remove('input-erro'); }); }

      btn.addEventListener('click', function(){
        limpaErros();
        msg.className = 'msg';
        msg.textContent = '';

        var grupoEl = document.querySelector('input[name="publico"]:checked');
        var erro = null, alvo = null;

        if(!nome.value.trim()){ erro = 'Digite seu nome.'; alvo = nome; }
        else if(!wpp.value.trim()){ erro = 'Digite seu WhatsApp.'; alvo = wpp; }
        else if(!emailValido(email.value.trim())){ erro = 'Digite um e-mail válido.'; alvo = email; }
        else if(!grupoEl){ erro = 'Selecione a qual grupo você pertence.'; }

        if(erro){
          msg.className = 'msg erro';
          msg.textContent = erro;
          if(alvo){ alvo.classList.add('input-erro'); alvo.focus(); }
          return;
        }

        var dados = {
          nome:  nome.value.trim(),
          whatsapp: wpp.value.trim(),
          email: email.value.trim(),
          grupo: grupoEl.value
        };

        if(URL_PLANILHA === "COLE_A_URL_AQUI"){
          msg.className = 'msg info';
          msg.textContent = '⚠️ Formulário ainda não conectado. Falta colar a URL da planilha (veja INSTRUCOES.md).';
          return;
        }

        btn.disabled = true;
        var textoOriginal = btn.textContent;
        btn.textContent = 'Enviando...';
        msg.className = 'msg info';
        msg.textContent = '';

        fetch(URL_PLANILHA, {
          method: 'POST',
          mode: 'no-cors',
          body: JSON.stringify(dados)
        }).then(function(){
          document.getElementById('okEmail').textContent = dados.email;
          form.style.display = 'none';
          document.getElementById('success').style.display = 'block';
          document.getElementById('success').scrollIntoView({behavior:'smooth', block:'center'});
        }).catch(function(){
          btn.disabled = false;
          btn.textContent = textoOriginal;
          msg.className = 'msg erro';
          msg.textContent = 'Não foi possível enviar agora. Tente novamente.';
        });
      });
    })();
