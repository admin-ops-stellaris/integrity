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
          // IMPORTANT: Capture the current handlers to avoid race conditions
          const capturedSuccessHandler = successHandler;
          const capturedFailureHandler = failureHandler;
          // Reset handlers after capturing so next chain starts fresh
          successHandler = null;
          failureHandler = null;
          
          return async (...args) => {
            try {
              const data = await apiCall(String(prop), args);
              if (typeof capturedSuccessHandler === "function") {
                capturedSuccessHandler(data);
              }
              return data;
            } catch (err) {
              console.error(err);
              if (typeof capturedFailureHandler === "function") capturedFailureHandler(err);
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
