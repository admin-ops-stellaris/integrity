// public/gas-shim.js
// Compatibility shim so existing google.script.run code can work in Node.
// It converts: google.script.run.withSuccessHandler(fn).someServerFunc(a,b)
// into: POST /api/someServerFunc  { args:[a,b] }

(function () {
  const apiCall = async (methodName, args) => {
    const res = await fetch(`/api/${methodName}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ args }),
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`API ${methodName} failed: ${res.status} ${text}`);
    }
    return await res.json();
  };

  function makeRunner() {
    let successHandler = null;
    let failureHandler = null;

    const runner = new Proxy(
      {},
      {
        get(_target, prop) {
          if (prop === "withSuccessHandler") {
            return (fn) => {
              successHandler = fn;
              return runner;
            };
          }
          if (prop === "withFailureHandler") {
            return (fn) => {
              failureHandler = fn;
              return runner;
            };
          }

          // Any other property is treated as a server function name
          return async (...args) => {
            try {
              const data = await apiCall(String(prop), args);
              if (typeof successHandler === "function") successHandler(data);
              return data;
            } catch (err) {
              console.error(err);
              if (typeof failureHandler === "function") failureHandler(err);
              throw err;
            }
          };
        },
      }
    );

    return runner;
  }

  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = makeRunner();
})();
