var ie = function() {
	var _obj, _doc;

	return {
		open: function(url, visible) {
			_obj = new ActiveXObject('InternetExplorer.Application');
			_obj.Visible = visible !== undefined ? visible : true;
			this.view(url);
		},
			view: function(url) {
				_obj.Navigate(url !== undefined || url != '' ? url : 'about:blank');
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
			getClass: function(cName) {
				try {
					return _doc.getElementsByClassName(cName);
				} catch (e) {
					for (var i = 0, j = 0, ret = [], tags = _doc.getElementsByTagName('*'); tag = tags[i++];) {
						if (tag.className == cName) {
							ret[j++] = tag;
						}
					}
					return ret;
				}
			},
			getTag: function(tName) {
				tName = tName || '*';
				return _doc.getElementsByTagName(tName);
			},
			getName: function(Name) {
				return _doc.getElementsByName(Name);
			}
	};
}();

