(() => {
  function switchOS(os) {
    if (os !== 'mac' && os !== 'win') {
      return;
    }

    document.querySelectorAll('.os-toggle button[data-os]').forEach((button) => {
      button.classList.toggle('active', button.dataset.os === os);
    });

    const macPanel = document.getElementById('panel-mac');
    const winPanel = document.getElementById('panel-win');
    if (macPanel) {
      macPanel.classList.toggle('active', os === 'mac');
    }
    if (winPanel) {
      winPanel.classList.toggle('active', os === 'win');
    }
  }

  function copyCmd(button) {
    if (!navigator.clipboard || typeof navigator.clipboard.writeText !== 'function') {
      return;
    }

    const terminal = button.closest('.terminal');
    const pre = terminal ? terminal.querySelector('pre') : null;
    if (!pre || !pre.textContent) {
      return;
    }

    const text = pre.textContent
      .split('\n')
      .filter((line) => !line.trim().startsWith('#'))
      .map((line) => line.replace(/^\$ /, ''))
      .filter((line) => line.trim())
      .join('\n');

    navigator.clipboard.writeText(text).then(() => {
      button.textContent = 'Copied!';
      button.classList.add('copied');
      setTimeout(() => {
        button.textContent = 'Copy';
        button.classList.remove('copied');
      }, 2000);
    });
  }

  document.querySelectorAll('.os-toggle button[data-os]').forEach((button) => {
    button.addEventListener('click', () => {
      const os = button.dataset.os;
      if (os) {
        switchOS(os);
      }
    });
  });

  document.querySelectorAll('.terminal-copy').forEach((button) => {
    button.addEventListener('click', () => copyCmd(button));
  });

  const observer = new IntersectionObserver(
    (entries) => entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
      }
    }),
    { threshold: 0.1, rootMargin: '0px 0px -40px 0px' },
  );

  document.querySelectorAll('.reveal').forEach((element) => {
    observer.observe(element);
  });
})();
