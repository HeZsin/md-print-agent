// ==UserScript==
// @name         明道云静默打印助手
// @namespace    md.silent-print
// @version      0.1.0
// @description  拦截明道云 Web 端打印动作，静默推送到本地打印代理(print_agent.py)，不再弹浏览器打印框；代理不可用时自动回退原生打印
// @author       安行
// @match        https://*.mingdao.com/*
// @match        http://*.mingdao.com/*
// @run-at       document-start
// @noframes
// @grant        GM_xmlhttpRequest
// @grant        GM_registerMenuCommand
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      127.0.0.1
// @connect      localhost
// ==/UserScript==

/* 私有部署请追加自己的域名（去掉注释并改成实际地址）：
// @match        https://hap.yourcompany.com/*
*/

(() => {
  'use strict';

  // ---------------- 配置 ----------------
  const CFG = {
    agent: GM_getValue('agent', 'http://127.0.0.1:8765'),
    printer: GM_getValue('printer', ''),        // 空 = 系统默认打印机
    silent: GM_getValue('silent', true),        // 拦截成功后不弹浏览器打印框
    debug: GM_getValue('debug', false),
  };
  const agentUrl = (path) => CFG.agent.replace(/\/+$/, '') + path;
  const log = (...a) => { if (CFG.debug) console.log('[静默打印]', ...a); };

  // ---------------- 基础工具（CSP 安全：全部走 CSSOM，不写 style 属性/字符串） ----------------
  function el(tag, styles, text) {
    const n = document.createElement(tag);
    if (styles) for (const [k, v] of Object.entries(styles)) n.style.setProperty(k, v);
    if (text != null) n.textContent = text;
    return n;
  }

  function toast(msg, ok = true) {
    if (!document.body) { console.log('[静默打印]', msg); return; }
    const host = el('div', {
      position: 'fixed', right: '20px', bottom: '20px', 'z-index': '2147483647',
      'font-family': 'system-ui,sans-serif', display: 'flex',
      'flex-direction': 'column', gap: '8px', 'pointer-events': 'none',
    });
    const card = el('div', {
      background: ok ? '#0f7a3d' : '#b35900', color: '#fff',
      padding: '10px 16px', 'border-radius': '8px', 'font-size': '14px',
      'box-shadow': '0 4px 12px rgba(0,0,0,.25)', 'max-width': '360px',
      transition: 'opacity .4s',
    }, msg);
    host.appendChild(card);
    document.body.appendChild(host);
    setTimeout(() => card.style.setProperty('opacity', '0'), 2800);
    setTimeout(() => host.remove(), 3400);
  }

  const blobToDataUrl = (blob) => new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(fr.result);
    fr.onerror = rej;
    fr.readAsDataURL(blob);
  });
  const blobToBase64 = async (blob) => (await blobToDataUrl(blob)).split(',')[1];

  // ---------------- 与代理通信 ----------------
  function sendJob(job) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: agentUrl('/print'),
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify(job),
        timeout: 30000,
        onload: (r) => {
          let resp = {};
          try { resp = JSON.parse(r.responseText); } catch { /* 非JSON响应 */ }
          if (r.status === 200 && resp.ok) {
            toast('已发送打印：' + (resp.printer || ''));
            resolve(resp);
          } else {
            toast('打印代理报错：' + (resp.error || 'HTTP ' + r.status), false);
            reject(new Error(resp.error || r.status));
          }
        },
        onerror: () => { toast('连不上本地打印代理，请确认 print_agent.py 已运行', false); reject(new Error('unreachable')); },
        ontimeout: () => { toast('打印代理超时', false); reject(new Error('timeout')); },
      });
    });
  }

  function pingAgent() {
    GM_xmlhttpRequest({
      method: 'GET', url: agentUrl('/status'), timeout: 3000,
      onload: (r) => { if (r.status === 200) log('代理在线', r.responseText); },
      onerror: () => toast('未检测到本地打印代理，本次会话打印将走浏览器原生弹窗', false),
    });
  }

  // ---------------- 内容捕获 ----------------
  // 把打印文档克隆为自包含 HTML（样式/图片全部内联，供代理离线渲染）
  async function captureHtml(doc) {
    const clone = doc.documentElement.cloneNode(true);
    clone.querySelectorAll('script,noscript').forEach((n) => n.remove());
    await Promise.all([...clone.querySelectorAll('link[rel~="stylesheet"][href]')].map(async (link) => {
      try {
        const css = await fetch(link.href, { credentials: 'include' }).then((r) => r.text());
        const style = doc.createElement('style');
        style.textContent = css;
        link.replaceWith(style);
      } catch { /* 跨域样式保留原链接，无头渲染时仍可尝试加载 */ }
    }));
    await Promise.all([...clone.querySelectorAll('img[src]')].map(async (img) => {
      try {
        if (img.src.startsWith('data:')) return;
        img.src = await blobToDataUrl(await fetch(img.src, { credentials: 'include' }).then((r) => r.blob()));
      } catch { /* 失败保留原地址 */ }
    }));
    // ponytail: CSS 内部相对 url()、跨域图片未内联——打印模板若出现缺图/缺样式再针对性补
    return '<!DOCTYPE html>\n' + clone.outerHTML;
  }

  async function printDocument(doc, label) {
    const html = await captureHtml(doc);
    log('捕获成功', label, '长度', html.length);
    return sendJob({ html, printer: CFG.printer || undefined, copies: 1 });
  }

  // blob:/data: 预览文档（常见于 PDF 预览）→ 按内容类型分发
  async function printUrl(url) {
    const blob = await fetch(url, { credentials: 'include' }).then((r) => r.blob());
    log('抓取', url.slice(0, 60), blob.type, blob.size + 'B');
    if (blob.type.includes('pdf')) {
      return sendJob({ pdfBase64: await blobToBase64(blob), printer: CFG.printer || undefined });
    }
    return sendJob({ html: await blob.text(), printer: CFG.printer || undefined });
  }

  // ---------------- 挂钩 window.print ----------------
  function hookPrint(win, label) {
    if (!win || win.__mdspHooked) return;
    try { win.__mdspHooked = true; } catch { return; }
    let original = null;
    try { original = win.print.bind(win); } catch { /* 受限窗口 */ }
    win.print = async function () {
      log('window.print 被调用：', label);
      if (CFG.silent) {
        try {
          await printDocument(win.document, label);
          return; // 静默成功，不再弹浏览器打印框
        } catch (e) {
          log('静默打印失败，回退浏览器打印：', e && e.message);
          toast('静默打印失败，已回退浏览器打印框', false);
        }
      }
      if (original) original();
    };
    log('已挂钩', label);
  }

  // ---------------- 监视打印 iframe / PDF 预览 ----------------
  let floatBtn = null;
  let lastPrintUrl = null;

  function ensureFloatBtn() {
    if (floatBtn || !document.body) return;
    floatBtn = el('button', {
      position: 'fixed', right: '20px', bottom: '64px', 'z-index': '2147483647',
      background: '#0f7a3d', color: '#fff', border: 'none', padding: '10px 16px',
      'border-radius': '8px', 'font-size': '14px', cursor: 'pointer',
      'font-family': 'system-ui,sans-serif', 'box-shadow': '0 4px 12px rgba(0,0,0,.25)',
    }, '静默打印此文档');
    floatBtn.addEventListener('click', async () => {
      if (!lastPrintUrl) return;
      floatBtn.disabled = true;
      try { await printUrl(lastPrintUrl); }
      catch (e) { toast('抓取文档失败：' + (e && e.message), false); }
      floatBtn.disabled = false;
    });
    document.body.appendChild(floatBtn);
  }

  function scanFrames() {
    let blobUrl = null;
    for (const f of document.querySelectorAll('iframe,embed')) {
      const url = f.src || f.data || '';
      try {
        if (/^blob:|^data:/i.test(url)) { blobUrl = url; continue; }
        if (/\.pdf(\?|#|$)/i.test(url) || (f.type || '').includes('pdf')) { blobUrl = url; continue; }
        if (f.contentWindow) hookPrint(f.contentWindow, 'iframe:' + (url.slice(0, 50) || 'inline'));
      } catch { /* 跨域 iframe，跳过 */ }
    }
    lastPrintUrl = blobUrl;
    if (blobUrl) ensureFloatBtn();
    else if (floatBtn) { floatBtn.remove(); floatBtn = null; }
  }

  function startWatching() {
    const mo = new MutationObserver(scanFrames);
    mo.observe(document.documentElement, { childList: true, subtree: true });
    setInterval(scanFrames, 1500); // ponytail: 轮询兜底；SPA 渲染漏检时再加针对性选择器
    scanFrames();
    pingAgent();
  }

  // 弹出窗口打印兜底：hook window.open 返回的窗口
  const origOpen = window.open ? window.open.bind(window) : null;
  if (origOpen) {
    window.open = function (...args) {
      const w = origOpen(...args);
      if (w) setTimeout(() => { try { hookPrint(w, 'popup'); } catch { /* 跨域 */ } }, 500);
      return w;
    };
  }

  // ---------------- 设置面板 ----------------
  function openSettings() {
    const old = document.getElementById('mdsp-panel');
    if (old) { old.remove(); return; }

    const panel = el('div', {
      position: 'fixed', top: '80px', right: '20px', 'z-index': '2147483647',
      width: '320px', background: '#fff', color: '#222', border: '1px solid #ddd',
      'border-radius': '12px', padding: '16px', 'font-family': 'system-ui,sans-serif',
      'font-size': '14px', 'box-shadow': '0 8px 24px rgba(0,0,0,.18)',
    });
    panel.id = 'mdsp-panel';

    const mkRow = (labelText, control) => {
      const row = el('div', { margin: '10px 0 0' });
      const lab = el('div', { 'font-size': '12px', color: '#666', 'margin-bottom': '4px' }, labelText);
      control.style.setProperty('width', '100%');
      control.style.setProperty('box-sizing', 'border-box');
      control.style.setProperty('padding', '6px 8px');
      control.style.setProperty('border', '1px solid #ccc');
      control.style.setProperty('border-radius', '6px');
      row.append(lab, control);
      return row;
    };

    const inAgent = el('input');
    inAgent.value = CFG.agent;
    const inPrinter = el('input');
    inPrinter.setAttribute('list', 'mdsp-printers');
    inPrinter.placeholder = '留空 = 系统默认打印机';
    inPrinter.value = CFG.printer;
    const dl = el('datalist');
    dl.id = 'mdsp-printers';

    const cbSilent = el('input');
    cbSilent.type = 'checkbox';
    cbSilent.checked = CFG.silent;
    const cbDebug = el('input');
    cbDebug.type = 'checkbox';
    cbDebug.checked = CFG.debug;

    const rowSilent = el('label', { display: 'flex', 'align-items': 'center', gap: '8px', margin: '12px 0 0', cursor: 'pointer' });
    rowSilent.append(cbSilent, el('span', null, '静默模式（不弹浏览器打印框）'));
    const rowDebug = el('label', { display: 'flex', 'align-items': 'center', gap: '8px', margin: '6px 0 0', cursor: 'pointer' });
    rowDebug.append(cbDebug, el('span', null, '调试日志（控制台输出）'));

    const btnRow = el('div', { display: 'flex', gap: '8px', margin: '14px 0 0' });
    const mkBtn = (text, primary) => {
      const b = el('button', {
        flex: '1', padding: '8px 0', border: 'none', 'border-radius': '6px',
        'font-size': '13px', cursor: 'pointer',
        background: primary ? '#0f7a3d' : '#eee', color: primary ? '#fff' : '#333',
      }, text);
      return b;
    };
    const btnLoad = mkBtn('获取打印机列表', false);
    const btnSave = mkBtn('保存', true);

    btnLoad.addEventListener('click', () => {
      GM_xmlhttpRequest({
        method: 'GET', url: agentUrl('/printers'), timeout: 5000,
        onload: (r) => {
          try {
            const resp = JSON.parse(r.responseText);
            dl.textContent = '';
            (resp.printers || []).forEach((p) => {
              const opt = el('option');
              opt.value = p.name;
              if (p.default) opt.label = '默认';
              dl.appendChild(opt);
            });
            toast('已获取 ' + (resp.printers || []).length + ' 台打印机，点击打印机输入框选择');
          } catch { toast('解析打印机列表失败', false); }
        },
        onerror: () => toast('连不上打印代理，请先运行 print_agent.py', false),
      });
    });
    btnSave.addEventListener('click', () => {
      CFG.agent = inAgent.value.trim() || 'http://127.0.0.1:8765';
      CFG.printer = inPrinter.value.trim();
      CFG.silent = cbSilent.checked;
      CFG.debug = cbDebug.checked;
      GM_setValue('agent', CFG.agent);
      GM_setValue('printer', CFG.printer);
      GM_setValue('silent', CFG.silent);
      GM_setValue('debug', CFG.debug);
      panel.remove();
      toast('设置已保存');
      pingAgent();
    });

    panel.append(
      el('div', { 'font-weight': '600', 'font-size': '15px' }, '静默打印设置'),
      mkRow('本地代理地址', inAgent),
      mkRow('打印机', inPrinter),
      dl,
      rowSilent,
      rowDebug,
      btnRow,
    );
    btnRow.append(btnLoad, btnSave);
    document.body.appendChild(panel);
  }

  // ---------------- 启动 ----------------
  GM_registerMenuCommand('设置（静默打印）', openSettings);
  GM_registerMenuCommand(CFG.debug ? '关闭调试日志' : '开启调试日志', () => {
    CFG.debug = !CFG.debug;
    GM_setValue('debug', CFG.debug);
    toast('调试日志已' + (CFG.debug ? '开启' : '关闭'));
  });

  hookPrint(window, 'top');
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startWatching);
  } else {
    startWatching();
  }
})();
