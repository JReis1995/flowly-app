window.google = {
  script: {
    run: (function() {
      let successHandler = null;
      let failureHandler = null;

      const runner = {
        withSuccessHandler(callback) {
          successHandler = callback;
          return proxy;
        },
        withFailureHandler(callback) {
          failureHandler = callback;
          return proxy;
        }
      };

      const proxy = new Proxy(runner, {
        get(target, prop) {
          if (prop in target) return target[prop];

          return (...args) => {
            const currentSuccess = successHandler;
            const currentFailure = failureHandler;

            // Reset handlers immediately for the next call chaining
            successHandler = null;
            failureHandler = null;

            if (typeof flowlyAPI !== 'undefined' && typeof flowlyAPI.call === 'function') {
              flowlyAPI.call(prop, args)
                .then(response => {
                  if (currentSuccess) currentSuccess(response);
                })
                .catch(error => {
                  if (currentFailure) currentFailure(error);
                  else console.error(`Error in GAS shim call [${prop}]:`, error);
                });
            } else {
              console.error('flowlyAPI or flowlyAPI.call is not defined. Ensure flowlyAPI is loaded before shim.js.');
            }
          };
        }
      });

      return proxy;
    })()
  }
};
