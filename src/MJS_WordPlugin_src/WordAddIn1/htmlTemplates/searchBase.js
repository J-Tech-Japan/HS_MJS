
// Configuration Constants
const SEARCH_CONFIG = {
  // Character conversion maps
  WIDE_CHARACTERS: [
    "０","１","２","３","４","５","６","７","８","９",
    "Ａ","Ｂ","Ｃ","Ｄ","Ｅ","Ｆ","Ｇ","Ｈ","Ｉ","Ｊ","Ｋ","Ｌ","Ｍ","Ｎ","Ｏ","Ｐ","Ｑ","Ｒ","Ｓ","Ｔ","Ｕ","Ｖ","Ｗ","Ｘ","Ｙ","Ｚ",
    "ａ","ｂ","ｃ","ｄ","ｅ","ｆ","ｇ","ｈ","ｉ","ｊ","ｋ","ｌ","ｍ","ｎ","ｏ","ｐ","ｑ","ｒ","ｓ","ｔ","ｕ","ｖ","ｗ","ｘ","ｙ","ｚ",
    "ガ","ギ","グ","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","ビ","ブ","ベ","ボ","パ","ピ","プ","ペ","ポ",
    "。","「","」","、","ヲ","ァ","ィ","ゥ","ェ","ォ","ャ","ュ","ョ","ッ","ー",
    "ア","イ","ウ","エ","オ","カ","キ","ク","ケ","コ","サ","シ","ス","セ","ソ","タ","チ","ツ","テ","ト","ナ","ニ","ヌ","ネ","ノ","ハ","ヒ","フ","ヘ","ホ","マ","ミ","ム","メ","モ","ヤ","ユ","ヨ","ラ","リ","ル","レ","ロ","ワ","ン"
  ],
  
  NARROW_CHARACTERS: [
    "0","1","2","3","4","5","6","7","8","9",
    "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
    "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
    "ｶﾞ","ｷﾞ","ｸﾞ","ｹﾞ","ｺﾞ","ｻﾞ","ｼﾞ","ｽﾞ","ｾﾞ","ｿﾞ","ﾀﾞ","ﾁﾞ","ﾂﾞ","ﾃﾞ","ﾄﾞ","ﾊﾞ","ﾋﾞ","ﾌﾞ","ﾍﾞ","ﾎﾞ","ﾊﾟ","ﾋﾟ","ﾌﾟ","ﾍﾟ","ﾎﾟ",
    "｡","｢","｣","､","ｦ","ｧ","ｨ","ｩ","ｪ","ｫ","ｬ","ｭ","ｮ","ｯ","ｰ",
    "ｱ","ｲ","ｳ","ｴ","ｵ","ｶ","ｷ","ｸ","ｹ","ｺ","ｻ","ｼ","ｽ","ｾ","ｿ","ﾀ","ﾁ","ﾂ","ﾃ","ﾄ","ﾅ","ﾆ","ﾇ","ﾈ","ﾉ","ﾊ","ﾋ","ﾌ","ﾍ","ﾎ","ﾏ","ﾐ","ﾑ","ﾒ","ﾓ","ﾔ","ﾕ","ﾖ","ﾗ","ﾘ","ﾙ","ﾚ","ﾛ","ﾜ","ﾝ"
  ],
  
  HIGHLIGHT_PATTERNS: [
    "(?:０|0)","(?:１|1)","(?:２|2)","(?:３|3)","(?:４|4)","(?:５|5)","(?:６|6)","(?:７|7)","(?:８|8)","(?:９|9)",
    "(?:Ａ|A|ａ|a)","(?:Ｂ|B|ｂ|b)","(?:Ｃ|C|ｃ|c)","(?:Ｄ|D|ｄ|d)","(?:Ｅ|E|ｅ|e)","(?:Ｆ|F|ｆ|f)","(?:Ｇ|G|ｇ|g)","(?:Ｈ|H|ｈ|h)","(?:Ｉ|I|ｉ|i)","(?:Ｊ|J|ｊ|j)","(?:Ｋ|K|ｋ|k)","(?:Ｌ|L|ｌ|l)","(?:Ｍ|M|ｍ|m)","(?:Ｎ|N|ｎ|n)","(?:Ｏ|O|ｏ|o)","(?:Ｐ|P|ｐ|p)","(?:Ｑ|Q|ｑ|q)","(?:Ｒ|R|ｒ|r)","(?:Ｓ|S|ｓ|s)","(?:Ｔ|T|ｔ|t)","(?:Ｕ|U|ｕ|u)","(?:Ｖ|V|ｖ|v)","(?:Ｗ|W|ｗ|w)","(?:Ｘ|X|ｘ|x)","(?:Ｙ|Y|ｙ|y)","(?:Ｚ|Z|ｚ|z)",
    "(?:ガ|ｶﾞ)","(?:ギ|ｷﾞ)","(?:グ|ｸﾞ)","(?:ゲ|ｹﾞ)","(?:ゴ|ｺﾞ)","(?:ザ|ｻﾞ)","(?:ジ|ｼﾞ)","(?:ズ|ｽﾞ)","(?:ゼ|ｾﾞ)","(?:ゾ|ｿﾞ)","(?:ダ|ﾀﾞ)","(?:ヂ|ﾁﾞ)","(?:ヅ|ﾂﾞ)","(?:デ|ﾃﾞ)","(?:ド|ﾄﾞ)","(?:バ|ﾊﾞ)","(?:ビ|ﾋﾞ)","(?:ブ|ﾌﾞ)","(?:ベ|ﾍﾞ)","(?:ボ|ﾎﾞ)","(?:パ|ﾊﾟ)","(?:ピ|ﾋﾟ)","(?:プ|ﾌﾟ)","(?:ペ|ﾍﾟ)","(?:ポ|ﾎﾟ)",
    "(?:。|｡)","(?:「|｢)","(?:」|｣)","(?:、|､)","(?:ヲ|ｦ)","(?:ァ|ｧ)","(?:ィ|ｨ)","(?:ゥ|ｩ)","(?:ェ|ｪ)","(?:ォ|ｫ)","(?:ャ|ｬ)","(?:ュ|ｭ)","(?:ョ|ｮ)","(?:ッ|ｯ)","(?:ー|ｰ)",
    "(?:ア|ｱ)","(?:イ|ｲ)","(?:ウ|ｳ)","(?:エ|ｴ)","(?:オ|ｵ)","(?:カ|ｶ)","(?:キ|ｷ)","(?:ク|ｸ)","(?:ケ|ｹ)","(?:コ|ｺ)","(?:サ|ｻ)","(?:シ|ｼ)","(?:ス|ｽ)","(?:セ|ｾ)","(?:ソ|ｿ)","(?:タ|ﾀ)","(?:チ|ﾁ)","(?:ツ|ﾂ)","(?:テ|ﾃ)","(?:ト|ﾄ)","(?:ナ|ﾅ)","(?:ニ|ﾆ)","(?:ヌ|ﾇ)","(?:ネ|ﾈ)","(?:ノ|ﾉ)","(?:ハ|ﾊ)","(?:ヒ|ﾋ)","(?:フ|ﾌ)","(?:ヘ|ﾍ)","(?:ホ|ﾎ)","(?:マ|ﾏ)","(?:ミ|ﾐ)","(?:ム|ﾑ)","(?:メ|ﾒ)","(?:モ|ﾓ)","(?:ヤ|ﾔ)","(?:ユ|ﾕ)","(?:ヨ|ﾖ)","(?:ラ|ﾗ)","(?:リ|ﾘ)","(?:ル|ﾙ)","(?:レ|ﾚ)","(?:ロ|ﾛ)","(?:ワ|ﾜ)","(?:ン|ﾝ)"
  ],
  
  // Messages
  MESSAGES: {
    SEARCH_HELP: "2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。",
    NO_RESULTS: "検索条件に一致するトピックはありません。"
  },
  
  // CSS Selectors
  SELECTORS: {
    SEARCH_FIELD: ".wSearchField",
    SEARCH_RESULTS_BLOCK: ".wSearchResultItemsBlock",
    SEARCH_RESULTS_END: ".wSearchResultsEnd",
    SEARCH_MSG: "#searchMsg",
    TOPIC_IFRAME: "iframe.topic",
    KEYWORD: ".keyword",
    TOC_BOOK: "ul.toc li.book",
    SEARCH_INPUT: ".search-input"
  },
  
  // CSS Classes
  CSS_CLASSES: {
    HIDDEN: "rh-hide",
    SEARCH_RESULT_ITEM: "wSearchResultItem",
    SEARCH_RESULT_TITLE: "wSearchResultTitle",
    SEARCH_CONTENT: "wSearchContent",
    SEARCH_CONTEXT: "wSearchContext",
    NO_LINK: "nolink"
  },
  
  // Styles
  HIGHLIGHT_STYLE: "color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);",
  
  // MutationObserver configuration
  MUTATION_OBSERVER: {
    DEBOUNCE_DELAY: 500, // milliseconds to wait before re-highlighting
    CONFIG: {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: false // We don't need to watch attribute changes
    }
  }
};

/**
 * WebHelp Search Class
 */
class WebHelpSearch {
  constructor() {
    this.searchWords = window.searchWords;
    this.currentSearchValue = ""; // Track current search value
    this.mutationObserver = null; // MutationObserver instance
    this.debounceTimer = null; // Debounce timer for DOM changes
    this.initializeEventHandlers();
  }

  /**
   * Initialize all event handlers
   */
  initializeEventHandlers() {
    // Table of contents book click handler
    $(document).on("click", SEARCH_CONFIG.SELECTORS.TOC_BOOK, this.handleTocBookClick.bind(this));
    
    // Reset search field handlers
    $(SEARCH_CONFIG.SELECTORS.SEARCH_FIELD).each(function() {
      $(this).off();
    });
    
    // Search field input handler
    $(document).on("keyup", SEARCH_CONFIG.SELECTORS.SEARCH_FIELD, this.handleSearchInput.bind(this));
    
    // Topic iframe load handler
    $(SEARCH_CONFIG.SELECTORS.TOPIC_IFRAME).on("load", this.handleTopicLoad.bind(this));
  }

  /**
   * Handle table of contents book click
   */
  handleTocBookClick(event) {
    const $this = $(event.currentTarget);
    if ($this.children("a[href='#'],a[href='javascript:void 0;']").length === 0) {
      $this.children("a").each(function() {
        location.href = location.href.replace(location.hash, "") + "#t=" + $(this).attr("href");
      });
    }
  }

  /**
   * Handle search input
   */
  handleSearchInput(event) {
    const searchValue = $(event.currentTarget).val();
    this.currentSearchValue = searchValue; // Store current search value
    
    if (searchValue === "") {
      this.clearSearchResults();
    } else {
      this.performSearch(searchValue);
    }
  }

  /**
   * Handle topic iframe load
   */
  handleTopicLoad() {
    const $searchInput = $(SEARCH_CONFIG.SELECTORS.SEARCH_INPUT, document);
    const $searchField = $(SEARCH_CONFIG.SELECTORS.SEARCH_FIELD, document);
    
    if ($searchInput.is(":not(.rh-hide)") && $searchField.val() !== "") {
      const searchValue = $searchField.val();
      this.currentSearchValue = searchValue; // Store current search value
      this.highlightSearchTerms(searchValue);
    }
    
    // Set up MutationObserver for iframe content changes
    this.setupMutationObserver();
  }

  /**
   * Clear search results and reset UI
   */
  clearSearchResults() {
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_BLOCK).html("");
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_END)
      .addClass(SEARCH_CONFIG.CSS_CLASSES.HIDDEN)
      .attr("hidden", "");
    $(SEARCH_CONFIG.SELECTORS.SEARCH_MSG).html(SEARCH_CONFIG.MESSAGES.SEARCH_HELP);
    this.removeKeywordHighlights();
    this.currentSearchValue = ""; // Clear current search value
    this.disconnectMutationObserver(); // Disconnect observer when clearing
  }

  /**
   * Perform search operation
   */
  performSearch(searchValue) {
    $(SEARCH_CONFIG.SELECTORS.SEARCH_MSG).html("");
    
    const processedSearchWords = this.processSearchInput(searchValue);
    const searchQuery = this.buildSearchQuery(processedSearchWords);
    const findItems = this.searchWords.find(".search_word" + searchQuery);
    
    if (findItems.length > 0) {
      this.displaySearchResults(findItems);
      this.highlightSearchTerms(searchValue);
    } else {
      this.displayNoResults();
    }
  }

  /**
   * Process search input: normalize spaces and convert characters
   */
  processSearchInput(searchValue) {
    let processed = searchValue
      .replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2")
      .trim()
      .toLowerCase();
    
    // Convert wide characters to narrow characters
    for (let i = 0; i < SEARCH_CONFIG.WIDE_CHARACTERS.length; i++) {
      processed = processed.split(SEARCH_CONFIG.WIDE_CHARACTERS[i]).join(SEARCH_CONFIG.NARROW_CHARACTERS[i]);
    }
    
    return processed.split(" ");
  }

  /**
   * Build jQuery selector query for search
   */
  buildSearchQuery(searchWords) {
    return searchWords.map(word => `:contains(${word})`).join("");
  }

  /**
   * Display search results
   */
  displaySearchResults(findItems) {
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_END)
      .removeClass(SEARCH_CONFIG.CSS_CLASSES.HIDDEN)
      .removeAttr("hidden");
    
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_BLOCK).html("");
    
    findItems.each((index, item) => {
      const $item = $(item);
      const $parent = $item.parent();
      const displayText = $parent.find(".displayText").text();
      const itemId = $parent.attr("id");
      const title = $parent.find(".search_title").html();
      
      const resultHtml = this.createSearchResultItem(itemId, title, displayText);
      $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_BLOCK).append(resultHtml);
    });
    
    this.removeKeywordHighlights();
  }

  /**
   * Create search result item HTML
   */
  createSearchResultItem(itemId, title, displayText) {
    return $(`
      <div class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_RESULT_ITEM}'>
        <a class='${SEARCH_CONFIG.CSS_CLASSES.NO_LINK}' href='./${itemId}.html'>
          <div class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_RESULT_TITLE}'>${title}</div>
        </a>
        <div class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_CONTENT}'>
          <span class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_CONTEXT}'>${displayText}</span>
        </div>
      </div>
    `);
  }

  /**
   * Display no results message
   */
  displayNoResults() {
    this.removeKeywordHighlights();
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_END)
      .addClass(SEARCH_CONFIG.CSS_CLASSES.HIDDEN)
      .attr("hidden", "");
    
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_BLOCK).html("");
    
    const noResultsHtml = $(`
      <div class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_RESULT_ITEM}'>
        <div class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_CONTENT}'>
          <span class='${SEARCH_CONFIG.CSS_CLASSES.SEARCH_CONTEXT}'>${SEARCH_CONFIG.MESSAGES.NO_RESULTS}</span>
        </div>
      </div>
    `);
    
    $(SEARCH_CONFIG.SELECTORS.SEARCH_RESULTS_BLOCK).append(noResultsHtml);
  }

  /**
   * Highlight search terms in the content
   */
  highlightSearchTerms(searchValue) {
    const normalizedSearchValue = this.normalizeSearchValue(searchValue);
    const searchWords = normalizedSearchValue.split(" ");
    const escapedSearchWords = searchWords.map(word => this.escapeForRegex(this.escapeHtmlChars(word)));
    
    const highlightPattern = this.createHighlightPattern(escapedSearchWords);
    this.applyHighlighting(highlightPattern);
  }

  /**
   * Normalize search value for highlighting
   */
  normalizeSearchValue(searchValue) {
    let normalized = searchValue.split("　").join(" ").trim();
    normalized = normalized.split("  ").join(" ");
    
    for (let i = 0; i < SEARCH_CONFIG.WIDE_CHARACTERS.length; i++) {
      normalized = normalized.replace(SEARCH_CONFIG.WIDE_CHARACTERS[i], SEARCH_CONFIG.NARROW_CHARACTERS[i]);
    }
    
    return normalized;
  }

  /**
   * Escape HTML characters
   */
  escapeHtmlChars(text) {
    return text.replace(/>/g, "&gt;").replace(/</g, "&lt;");
  }

  /**
   * Escape characters for regex
   */
  escapeForRegex(text) {
    return text.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
  }

  /**
   * Create highlight pattern for regex
   */
  createHighlightPattern(searchWords) {
    let highlightPattern = searchWords.join("|");
    
    for (let i = 0; i < SEARCH_CONFIG.HIGHLIGHT_PATTERNS.length; i++) {
      const regex = new RegExp(SEARCH_CONFIG.HIGHLIGHT_PATTERNS[i], "gm");
      highlightPattern = highlightPattern.replace(regex, SEARCH_CONFIG.HIGHLIGHT_PATTERNS[i]);
    }
    
    return highlightPattern;
  }

  /**
   * Apply highlighting to the iframe content
   */
  applyHighlighting(highlightPattern) {
    const $topicBody = $(SEARCH_CONFIG.SELECTORS.TOPIC_IFRAME).contents().find("body");
    
    // Define replacement patterns
    const replacements = [
      { pattern: /&nbsp;(?=[^<>]*<)/gm, replacement: "　" },
      { pattern: /&gt;(?=[^<>]*<)/gm, replacement: ">" },
      { pattern: /&lt;(?=[^<>]*<)/gm, replacement: "<" },
      { pattern: /&quot;(?=[^<>]*<)/gm, replacement: '"' },
      { pattern: /&amp;(?=[^<>]*<)/gm, replacement: "&" },
      { 
        pattern: new RegExp("(" + highlightPattern + ")(?=[^<>]*<)", "gm"), 
        replacement: `<font class='keyword' style='${SEARCH_CONFIG.HIGHLIGHT_STYLE}'>$1</font>` 
      }
    ];
    
    let content = $topicBody.html();
    replacements.forEach(({ pattern, replacement }) => {
      content = content.replace(pattern, replacement);
    });
    
    $topicBody.html(content);
  }

  /**
   * Remove keyword highlights from iframe content
   */
  removeKeywordHighlights() {
    $(SEARCH_CONFIG.SELECTORS.TOPIC_IFRAME).contents().find(SEARCH_CONFIG.SELECTORS.KEYWORD).each(function() {
      const childNodes = this.childNodes;
      for (let i = 0; i < childNodes.length; i++) {
        this.parentNode.insertBefore(childNodes[i], this);
      }
      $(this).remove();
    });
  }

  /**
   * Setup MutationObserver to watch for DOM changes in iframe
   */
  setupMutationObserver() {
    // Disconnect any existing observer
    this.disconnectMutationObserver();
    
    try {
      const $iframe = $(SEARCH_CONFIG.SELECTORS.TOPIC_IFRAME);
      if ($iframe.length === 0) return;
      
      const iframeDocument = $iframe[0].contentDocument || $iframe[0].contentWindow.document;
      if (!iframeDocument || !iframeDocument.body) return;
      
      // Create new MutationObserver
      this.mutationObserver = new MutationObserver((mutations) => {
        this.handleDomMutations(mutations);
      });
      
      // Start observing iframe body for changes
      this.mutationObserver.observe(
        iframeDocument.body,
        SEARCH_CONFIG.MUTATION_OBSERVER.CONFIG
      );
      
      console.debug("MutationObserver setup for iframe content");
    } catch (error) {
      console.warn("Failed to setup MutationObserver:", error);
    }
  }

  /**
   * Handle DOM mutations in iframe
   */
  handleDomMutations(mutations) {
    // Check if we have a current search value
    if (!this.currentSearchValue || this.currentSearchValue.trim() === "") {
      return;
    }
    
    // Check if any mutations affect text content or add/remove nodes
    const shouldReHighlight = mutations.some(mutation => {
      // Skip mutations that are just our own highlighting changes
      if (mutation.type === 'childList') {
        // Check if added nodes contain keyword elements (our own highlighting)
        const addedKeywords = Array.from(mutation.addedNodes).some(node => 
          node.nodeType === Node.ELEMENT_NODE && 
          (node.classList?.contains('keyword') || node.querySelector?.('.keyword'))
        );
        
        // Skip if it's just our highlighting being added
        if (addedKeywords && mutation.addedNodes.length === 1) {
          return false;
        }
        
        // Re-highlight if nodes were added/removed (but not just our highlighting)
        return mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0;
      }
      
      // Re-highlight on text changes
      return mutation.type === 'characterData';
    });
    
    if (shouldReHighlight) {
      this.debouncedReHighlight();
    }
  }

  /**
   * Debounced re-highlighting to avoid excessive calls
   */
  debouncedReHighlight() {
    // Clear existing timer
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
    }
    
    // Set new timer
    this.debounceTimer = setTimeout(() => {
      this.reHighlightAfterDomChange();
    }, SEARCH_CONFIG.MUTATION_OBSERVER.DEBOUNCE_DELAY);
  }

  /**
   * Re-highlight search terms after DOM changes
   */
  reHighlightAfterDomChange() {
    if (!this.currentSearchValue || this.currentSearchValue.trim() === "") {
      return;
    }
    
    try {
      // Temporarily disconnect observer to avoid infinite loops
      this.disconnectMutationObserver();
      
      // Remove existing highlights
      this.removeKeywordHighlights();
      
      // Re-apply highlighting
      this.highlightSearchTerms(this.currentSearchValue);
      
      // Re-setup observer
      setTimeout(() => {
        this.setupMutationObserver();
      }, 100); // Small delay to ensure DOM is stable
      
      console.debug("Re-highlighted search terms after DOM change");
    } catch (error) {
      console.warn("Failed to re-highlight after DOM change:", error);
      // Re-setup observer even if highlighting failed
      this.setupMutationObserver();
    }
  }

  /**
   * Disconnect MutationObserver
   */
  disconnectMutationObserver() {
    if (this.mutationObserver) {
      this.mutationObserver.disconnect();
      this.mutationObserver = null;
    }
    
    // Clear debounce timer
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
  }
}

// Legacy function for backward compatibility
function selectorEscape(val) {
  return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
}

// Initialize when document is ready
$(function() {
  // Initialize the search functionality
  const webHelpSearch = new WebHelpSearch();
});
