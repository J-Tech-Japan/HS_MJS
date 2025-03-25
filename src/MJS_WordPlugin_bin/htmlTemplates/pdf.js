var p_pdf;
if(typeof p_pdfClass === 'undefined') {
	var p_pdfClass = function(){
		this.init();
	}
}

p_pdfClass.prototype = {
	init: function () {
		this.ww = window.innerWidth;
		this._frame = '';
		this._size = 'small';
		this._cookie = '';
		this._pankuzuSize = 10;
		
		this.doc;
		this.pdf = false;
		this.frameUrl = '';
		this.pdfPageMax = 0;
		this.pdfPrevImgPos = 0;
		this.pdfPrevImgMove = 0;
		this.pdfPrevImgH = 0;
		this.pdfPageCurrent = 1;
	},
	ready: function(){
		//console.log('p_pdf ready');
		p_pdf.setNavi();
		p_pdf.set_iframeLoad();
	},
	initialize: function(){
	
	},
	loaded: function(){
	
	},
	setNavi: function(){
		$('#print_page').on('click', function(){
			p_pdf.openPdf();
			p_pdf.htmlcanvas();
		});
		
		$('#buttonCloseModalPdf, #buttonCancelPdf').on('click', function(){
			p_pdf.closePdf();
		});
		
		$('#buttonOutputPdf').on('click', function(){
			p_pdf.savePdf();
			p_pdf.closePdf();
		});
		
		$('#modalPdfPageNext').on('click', function(){
			if(!$(this).hasClass('off')){
				p_pdf.nextPdfPage();
			}
		});
		
		$('#modalPdfPagePrev').on('click', function(){
			if(!$(this).hasClass('off')){
				p_pdf.prevPdfPage();
			}
		});
	},
	openPdf: function(){
		$('#modalPdf').addClass('show');
	},
	closePdf: function(){
		$('#modalPdf').removeClass('show');
		p_pdf.resetPdf();
	},
	set_iframeLoad: function(){
		$('iframe').on('load', function(){
			p_pdf.frameUrl = (location.hash).replace('#t=', '');
			p_pdf.getHtml(p_pdf.frameUrl);
		});
	},
	getHtml: function(htmlName){
		$.ajax(htmlName, {
			timeout : 1000,
			datatype:'html'
		}).then(function(data){
			var out_html = data;
			//console.log(out_html);
			
			var html_array;
			
			//if(out_html.indexOf('<body style="text-justify-trim: punctuation;">') > 0){
			//	html_array = data.split('<body style="text-justify-trim: punctuation;">');
			//} else {
			//	html_array = data.split('<body>');
			//}
			html_array = data.split(/<body[^>]*>/);
			
			//console.log(html_array);
			
			var body_html = (html_array[1]).replace('</body>', '');
			body_html = (body_html).replace('</html>', '');
			//console.log(body_html);
			
			//$('#modalPdfLoader').append(body_html);
			
			//p_pdf.htmlcanvas();
			
		},function(jqXHR, textStatus) {
		});
	},
	htmlcanvas: function(){
		p_pdf.resetPdf();
		
		p_pdf.doc = new jsPDF("p", "mm", "a4");
		var _frame = document.getElementsByClassName("topic");
		var html = '';
		var nodeNum = 1;
		var nodeNum_main = 1;
		
		if(location.hash !== '') {
			_html = _frame[1].contentWindow.document.body.childNodes[nodeNum].innerHTML;
		} else {
			nodeNum = 0;
			_html = '<div>' + _frame[1].contentWindow.document.body.innerHTML + '</div>';
			_frame[1].contentWindow.document.body.innerHTML = _html;
		}
		
		//console.log(_html);
		
		//console.log(_frame[1].contentWindow.document);
		//console.log(_frame[1].contentWindow.document.head);
		//console.log(_frame[1].contentWindow.document.getElementsByTagName('title')[0].innerText);
		//console.log(_html);
		//console.log(_frame[1].contentWindow.document.body);
		//console.log(_frame[1].contentWindow.document.body.childNodes);
		//$(_frame[1].contentWindow.document.body.childNodes[nodeNum]).css("padding-left","50px");
		_frame[1].contentWindow.document.body.scrollTop = 0;
		_frame[1].contentWindow.document.documentElement.scrollTop = 0;
		_frame[1].contentWindow.document.body.classList.add('pdf');
		_frame[1].contentWindow.document.body.classList.remove('f_' + fc._size);
		_frame[1].contentWindow.document.body.classList.add('f_medium');
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.add('pdf');
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.remove('f_' + fc._size);
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.add('f_medium');
		//console.log(_frame[1].contentWindow.document.body.childNodes[nodeNum_main]);
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].childNodes[nodeNum].style.fontFamily = "'HiraginoSans-W2','ヒラギノ角ゴシック W2', 'メイリオ', Meiryo, 'ＭＳ Ｐゴシック', 'MS PGothic', sans-serif, Arial, Osaka";
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].childNodes[nodeNum].style.fontFeatureSettings = "'liga' 0";
		var _body = _frame[1].contentWindow.document.body.childNodes[nodeNum_main];
		var _div = $(_body).children('div').children('div');
		$(_div[0]).css('font-size', '11pt');
		
		$('body').append('<div id="pdfTitleArea">' + _frame[1].contentWindow.document.getElementsByTagName('title')[0].innerText + '</div>');
		$('#pdfTitleArea').css('width', $('.topic').width());
		
		//_frame[1].contentWindow.document.body.innerHTML += ('<div id="addPrintHtml" style="position: absolute; top: 0; z-index: -1; width: 1050px;">' + _html + '</div>');
		
		html2canvas($('#pdfTitleArea')).then(function(canvasTitle) {
			document.body.appendChild(canvasTitle);
			//_frame[1].contentWindow.document.body.childNodes[1]
			
			console.log(_frame[1].contentWindow.document.body.childNodes[1]);
			
			html2canvas(_frame[1].contentWindow.document.body.childNodes[nodeNum]).then(function (canvas) {
				//console.log(canvas);
				document.body.appendChild(canvas);
				
				var pageNum = 0;
				
				var titleImgData = canvasTitle.toDataURL('image/png');
				var titleWidth = 199;
				var titleHeight = canvasTitle.height * titleWidth / canvasTitle.width;
				
				var imgData = canvas.toDataURL('image/png');
				
				//$(_frame[1].contentWindow.document.body.childNodes[nodeNum]).css("padding-left","");
				
				//Page size in mm
				var pageHeight = 282;//max:297
				var imgWidth = 199;//210
				var imgHeight = canvas.height * (imgWidth / canvas.width);
				var heightLeft = imgHeight;
				var pdfImgHeight = 262;
				var position = 17.5;
				
				//console.log(imgHeight, heightLeft, (imgWidth / canvas.width));
				
				p_pdf.pdfPageMax++;
				p_pdf.doc.setFontSize(10);
				p_pdf.doc.addImage(imgData, 'PNG', 5.5, position, imgWidth, imgHeight);
				p_pdf.doc.setDrawColor(0);
				p_pdf.doc.setFillColor(255, 255, 255);
				p_pdf.doc.rect(0, 0, 210, 17.5, 'F');
				p_pdf.doc.rect(0, 279.5, 210, 17.5, 'F');
				p_pdf.doc.text(100, 285.5, String(p_pdf.pdfPageMax));
				p_pdf.doc.addImage(titleImgData, 'PNG', 0, 7.5, titleWidth, titleHeight);
				heightLeft -= pdfImgHeight;
				
				while (heightLeft >= 0) {
					p_pdf.pdfPageMax++;
					position = position - pdfImgHeight;
					p_pdf.doc.addPage();
					p_pdf.doc.setFontSize(10);
					p_pdf.doc.addImage(imgData, 'PNG', 5.5, position, imgWidth, imgHeight);
					p_pdf.doc.setDrawColor(0);
					p_pdf.doc.setFillColor(255, 255, 255);
					p_pdf.doc.rect(0, 0, 210, 17.5, 'F');
					p_pdf.doc.rect(0, 279.5, 210, 17.5, 'F');
					p_pdf.doc.text(100, 285.5, String(p_pdf.pdfPageMax));
					p_pdf.doc.addImage(titleImgData, 'PNG', 0, 7.5, titleWidth, titleHeight);
					heightLeft -= pdfImgHeight;
					
					//console.log(imgHeight, heightLeft, position, p_pdf.pdfPageMax);
				}
				
				//console.log(p_pdf.doc);
				
				var prevImgW = $('#modalPdfLoader').width();
				var prevImgH = (pdfImgHeight + 12) * (prevImgW / imgWidth);
				var prevPadding = 17.5 * (prevImgW / imgWidth);
				var prevTitlePadding = 7.5 * (prevImgW / imgWidth);
				
				//console.log((prevImgW / imgWidth));
				
				$('#modalPdfLoader').append('<span class="pageImg"><img src="' + imgData + '"></span>');
				$('#modalPdfLoader').append('<span class="title"><img src="' + titleImgData + '"></span>');
				$('#modalPdfLoader').find('.title').css({
					'padding-top': prevTitlePadding
				});
				$('#modalPdfLoader').find('.pageImg').css({
					'height': prevImgH, 'padding-top': prevPadding, 'margin-bottom': prevPadding,
				});
				
				p_pdf.pdfPrevImgPos = 0;
				p_pdf.pdfPrevImgMove = (prevImgH - prevPadding + 15.5);
				p_pdf.pdfPrevImgH = imgHeight * (prevImgW / imgWidth);
				
				$('#modalPdfPage').find('#modalPdfPageCurrent').html('1');
				$('#modalPdfPage').find('#modalPdfPageAll').html(p_pdf.pdfPageMax);
				
				if (p_pdf.pdfPageMax > 1) {
					$('#modalPdfPageNext').removeClass('off');
				}
				
				p_pdf.pdf = false;
			});
		});
	},
	resetPdf: function(){
		p_pdf.doc = '';
		p_pdf.pdfPageMax = 0;
		p_pdf.pdfPrevImgPos = 0;
		p_pdf.pdfPrevImgMove = 0;
		p_pdf.pdfPrevImgH = 0;
		p_pdf.pdfPageCurrent = 1;
		document.getElementById('modalPdfLoaderWrap').scrollTop = 0;
		
		if($('#pdfTitleArea').length > 0){
			$('#pdfTitleArea').remove();
		}
		
		if($('canvas').length > 0) {
			$('canvas').remove();
		}
		
		$('#modalPdfLoader').html('');
		$('#modalPdfPage').find('#modalPdfPageCurrent').html('-');
		$('#modalPdfPage').find('#modalPdfPageAll').html('-');
		$('#modalPdfPageNext, #modalPdfPagePrev').addClass('off');
		
		var _frame = document.getElementsByClassName("topic");
		var nodeNum_main = 1;
		_frame[1].contentWindow.document.body.classList.remove('pdf');
		_frame[1].contentWindow.document.body.classList.remove('f_medium');
		_frame[1].contentWindow.document.body.classList.add('f_' + fc._size);
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.remove('pdf');
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.remove('f_medium');
		_frame[1].contentWindow.document.body.childNodes[nodeNum_main].classList.add('f_' + fc._size);
		var _body = _frame[1].contentWindow.document.body;
		var _div = $(_body).children('div').children('div');
		$(_div[0]).css('font-size', fc._pankuzuSize + 'pt');
	},
	savePdf: function(){
		if(!p_pdf.pdf) {
			p_pdf.pdf = true;
			var pdfName = (p_pdf.frameUrl).replace('.html', '');
			
			// Fix: pdfName is empty in IE
			if(pdfName==""){
				pdfName = "pdf";
			}
			// End

			p_pdf.doc.save(pdfName + '.pdf');
		}
	},
	nextPdfPage: function(){
		p_pdf.pdfPrevImgPos += p_pdf.pdfPrevImgMove;
		p_pdf.pdfPageCurrent++;
		
		if(p_pdf.pdfPageCurrent > p_pdf.pdfPageMax){
			p_pdf.pdfPageCurrent = p_pdf.pdfPageMax;
		}
		
		document.getElementById('modalPdfLoaderWrap').scrollTop = 0;
		
		$('#modalPdfLoader').find('.pageImg').find('img').css('transform', 'translateY(-' + p_pdf.pdfPrevImgPos + 'px)');
		
		if(p_pdf.pdfPrevImgH < (p_pdf.pdfPrevImgPos + p_pdf.pdfPrevImgMove)){
			$('#modalPdfPageNext').addClass('off');
		} else {
			$('#modalPdfPageNext').removeClass('off');
		}
		
		if(p_pdf.pdfPrevImgPos > 0){
			$('#modalPdfPagePrev').removeClass('off');
		}
		
		$('#modalPdfPage').find('#modalPdfPageCurrent').html(p_pdf.pdfPageCurrent);
		
		
		// p_pdf.pdfPrevImgPos = 0;
		// p_pdf.pdfPrevImgMove = (prevImgH - prevPadding);
		// p_pdf.pdfPrevImgH = imgHeight;
	},
	prevPdfPage: function(){
		p_pdf.pdfPrevImgPos -= p_pdf.pdfPrevImgMove;
		p_pdf.pdfPageCurrent--;
		if(p_pdf.pdfPrevImgPos < 0){
			p_pdf.pdfPrevImgPos = 0;
		}
		
		if(p_pdf.pdfPageCurrent < 1){
			p_pdf.pdfPageCurrent = 1;
		}
		
		document.getElementById('modalPdfLoaderWrap').scrollTop = 0;
		
		$('#modalPdfLoader').find('.pageImg').find('img').css('transform', 'translateY(-' + p_pdf.pdfPrevImgPos + 'px)');
		
		if(p_pdf.pdfPrevImgPos <= 0){
			$('#modalPdfPagePrev').addClass('off');
		} else {
			$('#modalPdfPagePrev').removeClass('off');
		}
		
		if(p_pdf.pdfPrevImgH > (p_pdf.pdfPrevImgPos + p_pdf.pdfPrevImgMove)){
			$('#modalPdfPageNext').removeClass('off');
		}
		
		$('#modalPdfPage').find('#modalPdfPageCurrent').html(p_pdf.pdfPageCurrent);
	},
	closeToc: function(){
	
	},
	resizeHandler: function(){
	}
};

p_pdf = new p_pdfClass();

$(function(){
	p_pdf.ready();
	
	$(window).on('load', p_pdf.loaded);
});