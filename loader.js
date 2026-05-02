function cargarScript(path, type = "module") {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.type = type;
    script.src = path + "?v=" + window.APP_VERSION;

    script.onload = resolve;
    script.onerror = reject;

    document.body.appendChild(script);
  });
}
