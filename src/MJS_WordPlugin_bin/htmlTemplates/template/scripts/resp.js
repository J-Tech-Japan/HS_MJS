var respFrame;
if(typeof respFrameClass === 'undefined') {
	var respFrameClass = function(){
		this.init();
	}
}

respFrameClass.prototype = {
	init: function () {
		this.ww = window.innerWidth;
		this._frame = '';
		this.reloaded = false;
	},
	ready: function(){
		$('.MJS_bodyPict > img, .MJS_flowPict > img, .MJS_tdPict > img, .MJS_columnPict > img').attr('onclick', 'respFrame.openModal(event);');
		$('.MJS_bodyPict > img, .MJS_flowPict > img, .MJS_tdPict > img, .MJS_columnPict > img').wrap('<span class="icon_zoom"></span>');
		$('table').wrap('<div class="table-wrapper"></div>');
	},
	loaded: function(){
	},
	openModal: function(e){
		if($('body').hasClass('media-mobile')){
			var _img = e.target.src;
			respFrame.addModal(_img);
		}
	},
	addModal: function(img){
		var _html = '<div id="modalImg">' +
				'<div id="modalImgContents">' +
				'<div id="modalImgWrapper">' +
				'</div>' +
				'<div id="modalImgClose">閉じる</div>' +
				'</div>' +
				'<div id="modalImgBg"></div>' +
				'</div>';
		
		$('body').append(_html);
		$('#modalImgWrapper').html('<img src="' + img + '">');
		$('#modalImgClose').on('click', function(){
			$('#modalImg').remove();
		});
	}
}

respFrame = new respFrameClass();

$(function(){
	respFrame.ready();
	
	$(window).on('load', respFrame.loaded);
});