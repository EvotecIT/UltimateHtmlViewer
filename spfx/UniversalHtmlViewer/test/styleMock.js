module.exports = new Proxy(
  {},
  {
    get: function getStyleKey(_target, key) {
      if (key === '__esModule') {
        return false;
      }
      return key;
    },
  },
);
