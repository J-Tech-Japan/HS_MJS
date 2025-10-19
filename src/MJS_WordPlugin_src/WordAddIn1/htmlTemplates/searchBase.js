
var wide = Array("０","１","２","３","４","５","６","７","８","９","Ａ","Ｂ","Ｃ","Ｄ","Ｅ","Ｆ","Ｇ","Ｈ","Ｉ","Ｊ","Ｋ","Ｌ","Ｍ","Ｎ","Ｏ","Ｐ","Ｑ","Ｒ","Ｓ","Ｔ","Ｕ","Ｖ","Ｗ","Ｘ","Ｙ","Ｚ","ａ","ｂ","ｃ","ｄ","ｅ","ｆ","ｇ","ｈ","ｉ","ｊ","ｋ","ｌ","ｍ","ｎ","ｏ","ｐ","ｑ","ｒ","ｓ","ｔ","ｕ","ｖ","ｗ","ｘ","ｙ","ｚ","ガ","ギ","グ","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","ビ","ブ","ベ","ボ","パ","ピ","プ","ペ","ポ","。","「","」","、","ヲ","ァ","ィ","ゥ","ェ","ォ","ャ","ュ","ョ","ッ","ー","ア","イ","ウ","エ","オ","カ","キ","ク","ケ","コ","サ","シ","ス","セ","ソ","タ","チ","ツ","テ","ト","ナ","ニ","ヌ","ネ","ノ","ハ","ヒ","フ","ヘ","ホ","マ","ミ","ム","メ","モ","ヤ","ユ","ヨ","ラ","リ","ル","レ","ロ","ワ","ン");
var narrow = Array("0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","ｶﾞ","ｷﾞ","ｸﾞ","ｹﾞ","ｺﾞ","ｻﾞ","ｼﾞ","ｽﾞ","ｾﾞ","ｿﾞ","ﾀﾞ","ﾁﾞ","ﾂﾞ","ﾃﾞ","ﾄﾞ","ﾊﾞ","ﾋﾞ","ﾌﾞ","ﾍﾞ","ﾎﾞ","ﾊﾟ","ﾋﾟ","ﾌﾟ","ﾍﾟ","ﾎﾟ","｡","｢","｣","､","ｦ","ｧ","ｨ","ｩ","ｪ","ｫ","ｬ","ｭ","ｮ","ｯ","ｰ","ｱ","ｲ","ｳ","ｴ","ｵ","ｶ","ｷ","ｸ","ｹ","ｺ","ｻ","ｼ","ｽ","ｾ","ｿ","ﾀ","ﾁ","ﾂ","ﾃ","ﾄ","ﾅ","ﾆ","ﾇ","ﾈ","ﾉ","ﾊ","ﾋ","ﾌ","ﾍ","ﾎ","ﾏ","ﾐ","ﾑ","ﾒ","ﾓ","ﾔ","ﾕ","ﾖ","ﾗ","ﾘ","ﾙ","ﾚ","ﾛ","ﾜ","ﾝ");
var hilight = Array("(?:０|0)","(?:１|1)","(?:２|2)","(?:３|3)","(?:４|4)","(?:５|5)","(?:６|6)","(?:７|7)","(?:８|8)","(?:９|9)","(?:Ａ|A|ａ|a)","(?:Ｂ|B|ｂ|b)","(?:Ｃ|C|ｃ|c)","(?:Ｄ|D|ｄ|d)","(?:Ｅ|E|ｅ|e)","(?:Ｆ|F|ｆ|f)","(?:Ｇ|G|ｇ|g)","(?:Ｈ|H|ｈ|h)","(?:Ｉ|I|ｉ|i)","(?:Ｊ|J|ｊ|j)","(?:Ｋ|K|ｋ|k)","(?:Ｌ|L|ｌ|l)","(?:Ｍ|M|ｍ|m)","(?:Ｎ|N|ｎ|n)","(?:Ｏ|O|ｏ|o)","(?:Ｐ|P|ｐ|p)","(?:Ｑ|Q|ｑ|q)","(?:Ｒ|R|ｒ|r)","(?:Ｓ|S|ｓ|s)","(?:Ｔ|T|ｔ|t)","(?:Ｕ|U|ｕ|u)","(?:Ｖ|V|ｖ|v)","(?:Ｗ|W|ｗ|w)","(?:Ｘ|X|ｘ|x)","(?:Ｙ|Y|ｙ|y)","(?:Ｚ|Z|ｚ|z)","(?:ガ|ｶﾞ)","(?:ギ|ｷﾞ)","(?:グ|ｸﾞ)","(?:ゲ|ｹﾞ)","(?:ゴ|ｺﾞ)","(?:ザ|ｻﾞ)","(?:ジ|ｼﾞ)","(?:ズ|ｽﾞ)","(?:ゼ|ｾﾞ)","(?:ゾ|ｿﾞ)","(?:ダ|ﾀﾞ)","(?:ヂ|ﾁﾞ)","(?:ヅ|ﾂﾞ)","(?:デ|ﾃﾞ)","(?:ド|ﾄﾞ)","(?:バ|ﾊﾞ)","(?:ビ|ﾋﾞ)","(?:ブ|ﾌﾞ)","(?:ベ|ﾍﾞ)","(?:ボ|ﾎﾞ)","(?:パ|ﾊﾟ)","(?:ピ|ﾋﾟ)","(?:プ|ﾌﾟ)","(?:ペ|ﾍﾟ)","(?:ポ|ﾎﾟ)","(?:。|｡)","(?:「|｢)","(?:」|｣)","(?:、|､)","(?:ヲ|ｦ)","(?:ァ|ｧ)","(?:ィ|ｨ)","(?:ゥ|ｩ)","(?:ェ|ｪ)","(?:ォ|ｫ)","(?:ャ|ｬ)","(?:ュ|ｭ)","(?:ョ|ｮ)","(?:ッ|ｯ)","(?:ー|ｰ)","(?:ア|ｱ)","(?:イ|ｲ)","(?:ウ|ｳ)","(?:エ|ｴ)","(?:オ|ｵ)","(?:カ|ｶ)","(?:キ|ｷ)","(?:ク|ｸ)","(?:ケ|ｹ)","(?:コ|ｺ)","(?:サ|ｻ)","(?:シ|ｼ)","(?:ス|ｽ)","(?:セ|ｾ)","(?:ソ|ｿ)","(?:タ|ﾀ)","(?:チ|ﾁ)","(?:ツ|ﾂ)","(?:テ|ﾃ)","(?:ト|ﾄ)","(?:ナ|ﾅ)","(?:ニ|ﾆ)","(?:ヌ|ﾇ)","(?:ネ|ﾈ)","(?:ノ|ﾉ)","(?:ハ|ﾊ)","(?:ヒ|ﾋ)","(?:フ|ﾌ)","(?:ヘ|ﾍ)","(?:ホ|ﾎ)","(?:マ|ﾏ)","(?:ミ|ﾐ)","(?:ム|ﾑ)","(?:メ|ﾒ)","(?:モ|ﾓ)","(?:ヤ|ﾔ)","(?:ユ|ﾕ)","(?:ヨ|ﾖ)","(?:ラ|ﾗ)","(?:リ|ﾘ)","(?:ル|ﾙ)","(?:レ|ﾚ)","(?:ロ|ﾛ)","(?:ワ|ﾜ)","(?:ン|ﾝ)");

// Global variables for MutationObserver feature
var currentSearchValue = ""; // Current search keyword
var mutationObserver = null; // MutationObserver instance
var debounceTimer = null; // Debounce timer for DOM changes

function selectorEscape(val){
  return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
}

// Apply highlight to iframe content
function applyHighlight(searchValue) {
  var searchWordTmp = searchValue.split("　").join(" ").trim();
  searchWordTmp = searchWordTmp.split("  ").join(" ");
  for(var i = 0; i < wide.length; i++) {
    searchWordTmp = searchWordTmp.replace(wide[i], narrow[i]);
  }
  var searchWord = searchWordTmp.split(" ");
  for(var i = 0; i < searchWord.length; i++) {
    searchWord[i] = selectorEscape(searchWord[i].replace(">", "&gt;").replace("<", "&lt;"));
  }
  var hilightWord = searchWord.join("|");
  for(var i = 0; i < hilight.length; i++) {
    var reg = new RegExp(hilight[i], "gm");
    hilightWord = hilightWord.replace(reg, hilight[i]);
  }

  var reg = new RegExp("("+hilightWord+")(?=[^<>]*<)", "gm");
  var regnbsp = new RegExp("&nbsp;(?=[^<>]*<)", "gm");
  var reggt = new RegExp("&gt;(?=[^<>]*<)", "gm");
  var reglt = new RegExp("&lt;(?=[^<>]*<)", "gm");
  var regquot = new RegExp("&quot;(?=[^<>]*<)", "gm");
  var regamp = new RegExp("&amp;(?=[^<>]*<)", "gm");
  $("iframe.topic").contents().find("body").html($("iframe.topic").contents().find("body").html().replace(regnbsp, "　").replace(reggt, ">").replace(reglt, "<").replace(regquot, '"').replace(regamp, "&").replace(reg, "<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>"));
}

// Remove keyword highlights
function removeHighlight() {
  $("iframe.topic").contents().find(".keyword").each(function() {
    for(var i = 0; i < $(this)[0].childNodes.length; i++) {
      this.parentNode.insertBefore($(this)[0].childNodes[i], this);
    }
    $(this).remove();
  });
}

// Setup MutationObserver for iframe content
function setupMutationObserver() {
  disconnectMutationObserver();
  
  try {
    var $iframe = $("iframe.topic");
    if ($iframe.length === 0) return;
    
    var iframeDocument = $iframe[0].contentDocument || $iframe[0].contentWindow.document;
    if (!iframeDocument || !iframeDocument.body) return;
    
    mutationObserver = new MutationObserver(function(mutations) {
      if (!currentSearchValue || currentSearchValue.trim() === "") {
        return;
      }
      
      var shouldReHighlight = false;
      for (var i = 0; i < mutations.length; i++) {
        var mutation = mutations[i];
        if (mutation.type === 'childList') {
          var addedKeywords = false;
          for (var j = 0; j < mutation.addedNodes.length; j++) {
            var node = mutation.addedNodes[j];
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList && node.classList.contains('keyword')) {
                addedKeywords = true;
                break;
              }
              if (node.querySelector && node.querySelector('.keyword')) {
                addedKeywords = true;
                break;
              }
            }
          }
          
          if (addedKeywords && mutation.addedNodes.length === 1) {
            continue;
          }
          
          if (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0) {
            shouldReHighlight = true;
            break;
          }
        } else if (mutation.type === 'characterData') {
          shouldReHighlight = true;
          break;
        }
      }
      
      if (shouldReHighlight) {
        debouncedReHighlight();
      }
    });
    
    mutationObserver.observe(iframeDocument.body, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: false
    });
    
    console.debug("MutationObserver setup for iframe content");
  } catch (error) {
    console.warn("Failed to setup MutationObserver:", error);
  }
}

// Debounced re-highlighting
function debouncedReHighlight() {
  if (debounceTimer) {
    clearTimeout(debounceTimer);
  }
  
  debounceTimer = setTimeout(function() {
    reHighlightAfterDomChange();
  }, 500);
}

// Re-highlight after DOM change
function reHighlightAfterDomChange() {
  if (!currentSearchValue || currentSearchValue.trim() === "") {
    return;
  }
  
  try {
    disconnectMutationObserver();
    removeHighlight();
    applyHighlight(currentSearchValue);
    
    setTimeout(function() {
      setupMutationObserver();
    }, 100);
    
    console.debug("Re-highlighted search terms after DOM change");
  } catch (error) {
    console.warn("Failed to re-highlight after DOM change:", error);
    setupMutationObserver();
  }
}

// Disconnect MutationObserver
function disconnectMutationObserver() {
  if (mutationObserver) {
    mutationObserver.disconnect();
    mutationObserver = null;
  }
  
  if (debounceTimer) {
    clearTimeout(debounceTimer);
    debounceTimer = null;
  }
}

$(function(){
  $(document).on("click", "ul.toc li.book", function() {
    if($(this).children("a[href='#'],a[href='javascript:void 0;']").length == 0)
    {
      $(this).children("a").each(function(){
        location.href=location.href.replace(location.hash,"")+"#t="+$(this).attr("href");
      });
    }
  });
  $(".wSearchField").each(function() {
    $(this).off();
  });
  $(document).on("keyup", ".wSearchField", function(){
    if($(this).val() == "")
    {
      $(".wSearchResultItemsBlock").html("");
      $(".wSearchResultsEnd").addClass("rh-hide");
      $(".wSearchResultsEnd").attr("hidden", "");
      $("#searchMsg").html("2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。");
      removeHighlight();
      currentSearchValue = ""; // Clear current search value
      disconnectMutationObserver(); // Disconnect observer when clearing
    }
    else
    {
      $("#searchMsg").html("");
      currentSearchValue = $(this).val(); // Store current search value
      var searchWordTmp = $(this).val().replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2").trim().toLowerCase();
      for(i = 0; i < wide.length; i ++)
      {
        searchWordTmp = searchWordTmp.split(wide[i]).join(narrow[i]);
      }
      var searchWord = searchWordTmp.split(" ");
      var searchQuery = "";
      for(i = 0; i < searchWord.length; i ++)
      {
        searchQuery += ":contains(" + searchWord[i] + ")";
      }
      
      var findItems = searchWords.find(".search_word"+searchQuery);
      if(findItems.length != 0)
      {
        $(".wSearchResultsEnd").removeClass("rh-hide");
        $(".wSearchResultsEnd").removeAttr("hidden");
        $(".wSearchResultItemsBlock").html("");
        findItems.each(function() {
          var displayText = $(this).parent().find(".displayText").text();
          $(".wSearchResultItemsBlock").append($("<div class='wSearchResultItem'><a class='nolink' href='./"+$(this).parent().attr("id")+".html'><div class='wSearchResultTitle'>"+$(this).parent().find(".search_title").html()+"</div></a><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
        });
        removeHighlight();
        applyHighlight(currentSearchValue);
      }
      else
      {
        removeHighlight();
        $(".wSearchResultsEnd").addClass("rh-hide");
        $(".wSearchResultsEnd").attr("hidden", "");
        $(".wSearchResultItemsBlock").html("");
        displayText = "検索条件に一致するトピックはありません。";
        $(".wSearchResultItemsBlock").append($("<div class='wSearchResultItem'><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
      }
    }
  });
  $("iframe.topic").on("load", function(){
    if($(".search-input", document).is(":not(.rh-hide)") && ($(".wSearchField", document).val() != ""))
    {
      var searchValue = $(".wSearchField", document).val();
      currentSearchValue = searchValue; // Store current search value
      applyHighlight(searchValue);
    }
    
    // Setup MutationObserver for iframe content changes
    setupMutationObserver();
  });
});
