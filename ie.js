var ie = function() {
  var _obj = null, _doc = null;

  return {
    open: function(url, visible) {
      _obj = new ActiveXObject('InternetExplorer.Application');
      _obj.Visible = visible !== undefined ? visible : true;
      this.view(url);
    },
    view: function(url) {
      _obj.Navigate(url !== undefined ? url : 'about:blank');
      this.wait();
      _doc = _obj.document;
    },
    wait: function() {
      var start_date = new Date();
      while (_obj.Busy || _obj.ReadyState != 4) {
        if (typeof WScript === 'object') {
          WScript.Sleep(500);
        }
        if (new Date() - start_date >= 90 * 1000) {
          break;
        }
      }
    },
    close: function() {
      _obj.Quit();
    },
    getId: function(id) {
      return _doc.getElementById(id);
    },
    getClass: function(name) {
      try {
        return _doc.getElementsByClassName(name);
      } catch (e) {
        for (var i = 0, j = 0, ret = [], tags = _doc.getElementsByTagName('*');
              tag = tags[i++];) {
          if (tag.className == name) {
            ret[j++] = tag;
          }
        }
        return ret;
      }
    },
    getTag: function(tag) {
      tag = tag || '*';
      return _doc.getElementsByTagName(tag);
    },
    getName: function(name) {
      return _doc.getElementsByName(name);
    }
  };
}();

