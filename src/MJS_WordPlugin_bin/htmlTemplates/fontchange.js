var fc;
if(typeof fcClass === 'undefined') {
	var fcClass = function(){
		this.init();
	}
}

fcClass.prototype = {
	init: function () {
		this._frame = '';
		this._size = 'small';
		this._cookie = '';
		this._pankuzuSize = 10;
	},
	ready: function(){
		$('.fontsize_change').on('click', function(){
			fc.fontChange($(this));
		});
		
		fc._cookie = $.cookie('fontsize');
		if(typeof fc._cookie == "undefined"){
			fc.saveCookie();
		} else {
			fc._size = fc._cookie;
		}
		
		fc.fontChangeIndex();
		
		//console.log(fc._size, fc._cookie);
		
		fc._frame = document.getElementsByClassName("topic");
		$('.topic').on('load', function(){
			fc.frameLoaded();
		});
	},
	loaded: function(){
	},
	fontChange: function(obj){
		var id = obj.attr('id');
		fc._size = id.replace('fontsize_', '');
		fc._frame = document.getElementsByClassName("topic");
		console.log(id, fc._size);
		console.log(fc._frame[1]);
		console.log(fc._frame[1].contentWindow.document.body);
		
		$('.fontsize_change.active').removeClass('active');
		obj.addClass('active');
		
		fc.fontChangeIndex();
		fc.fontChangeIframe();
		fc.saveCookie();
	},
	fontChangeIndex: function(){
		$('body').removeClass('f_small');
		$('body').removeClass('f_medium');
		$('body').removeClass('f_large');
		$('body').addClass('f_' + fc._size);
	},
	fontChangeIframe: function(){
		var _body = fc._frame[1].contentWindow.document.body;
		$(_body).attr('class', '');
		$(_body).addClass('f_' + fc._size);
		
		fc.setPankuzuSize();
		var _div = $(_body).children('div').children('div');
		$(_div[0]).css('font-size', fc._pankuzuSize + 'pt');
	},
	setPankuzuSize: function(){
		switch(fc._size){
			case 'small': fc._pankuzuSize = 10; break;
			case 'medium': fc._pankuzuSize = 11; break;
			case 'large': fc._pankuzuSize = 12; break;
			default: fc._pankuzuSize = 10; break;
		}
	},
	saveCookie: function(){
		fc._cookie = fc._size;
		$.cookie('fontsize', fc._cookie, {expires: 365});
	},
	frameLoaded: function(){
		fc.fontChangeIframe();
		
		$('.fontsize_change.active').removeClass('active');
		$('#fontsize_' + fc._size).addClass('active');
	}
};

fc = new fcClass();

$(function(){
	fc.ready();
});