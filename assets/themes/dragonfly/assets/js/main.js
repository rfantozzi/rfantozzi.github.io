/*
* Theme Name: Dragonfly
* Theme URI: http://previews.sunflowertheme.com/dragonfly
* Description: Complete Theme For Ghost
* Author: Sunflowertheme
* Author URI: http://www.sunflowertheme.com
* Version:1.4
*/

/*  TABLE OF CONTENT
    1. Common function
    2. Initialing
*/
/*================================================================*/
/*  1. Common function
/*================================================================*/
function path(){
    var args = arguments,result = [];
    for(var i = 0; i < args.length; i++)
        result.push(args[i].replace('@', '/assets/js/syntaxhighlighter/'));
    return result;
};
var ghostApp={
    formatWidth: function(){
        if($('#isotope-content').length){
            var contentWidth = $('#isotope-content').width();
            var isGallery=false;
            var columns = 3;
            var columnPadding=20;
            if($('body').is('.gallery')){
                isGallery=true;
                columnPadding=0;
                if (contentWidth<=480){
                    columns=1;
                }
                else if (contentWidth<=640){
                    columns=2;
                }
                else if (contentWidth<=800){
                    columns=3;
                }
                else if(contentWidth<=1140){
                    columns=4;
                }
                else if(contentWidth<=1380){
                    columns=5;
                }
                else if(contentWidth<=1620){
                    columns=6;
                }
                else if(contentWidth>1620){
                    columns=7;
                }
            }
            else{
                if (contentWidth<640){
                    columns=1;
                }
                else if (contentWidth<=966){
                    columns=2;
                }
                else if(contentWidth<=1140){
                    columns=3;
                }
                else if(contentWidth<=1620){
                    columns=4;
                }
                else if(contentWidth>1620){
                    columns=5;
                }
            }
            var itemWidth = Math.floor((contentWidth-(columnPadding*(columns+1)))/columns);
            $('.item-list').each(function(){
                if(isGallery){
                    var ratio='';
                    if($(this).data('ratio'))
                        ratio=$(this).data('ratio');
                    if(ratio!='' && columns>2){
                        var ratioData=ratio.split('x');
                        if(ratioData.length==2){
                            var newWidth=Math.floor(itemWidth*parseInt(ratioData[0]));
                            var newHeight=Math.floor(itemWidth*parseInt(ratioData[1]));
                            $(this).css({"width":newWidth+"px"});
                            $(this).css({"height":newHeight+"px"});
                        }
                        else{
                            $(this).css({"width":itemWidth+"px"});
                            $(this).css({"height":itemWidth+"px"});
                        }
                    }
                    else{
                        $(this).css({"width":itemWidth+"px"});
                        $(this).css({"height":itemWidth+"px"});
                    }
                }
                else{
                    $(this).css({"width":itemWidth+"px"});
                }
            });
        }
    },
    responsiveImages:function(){
        $('.post').each(function(){
            var $imgList=$(this).find('img');
            if($imgList.length){
                $imgList.each(function(index, el) {
                    var alt=$(this).attr('alt');     
                    $(this).addClass('img-responsive');               
                    if(alt.indexOf('full-col')>=0){
                        $(this).addClass('full-col');
                    }
                    if(alt.indexOf('align-left')>=0){
                        $(this).addClass('align-left');
                    }
                    if(alt.indexOf('align-right')>=0){
                        $(this).addClass('align-right');
                    }
                    if(alt.indexOf('align-center')>=0){
                        $(this).addClass('align-center');
                    }
                    if(alt.indexOf('no-responsive')>=0 && !$('body').is('.home-template')){
                        $(this).css('width','auto');
                        $(this).removeClass('img-responsive');
                    }
                });
            }
        });        
    },
    centerPostDesc:function(){
        $('.post .wrap').each(function() {
            var $innerContent=$('.inner-content',$(this));
            var innerContentMargin=Math.floor($innerContent.height()/2);
            $innerContent.css('margin-top', '-'+innerContentMargin+'px');
        });
    },
    reformatPost:function(){
        if($('.post').length){
            var isGallery=false;
            if($('body').is('.gallery'))
                isGallery=true;
            $('.post .wrap:not(.formated)').each(function() {
                if($(this).find('.post-excerpt').has('img').length){
                    var $excerpt=$('.post-excerpt',$(this));
                    var $postImg=$excerpt.find('img');
                    if(isGallery){
                        $excerpt.css('background-image', 'url("'+$postImg.attr('src')+'")');
                        var regExp=/ratio-(\d)x(\d)/;
                        var ratio ='';
                        var regResult= $postImg.attr('alt').match(regExp);
                        if(regResult!=null){
                            $(this).parent().data('ratio',regResult[1]+'x'+regResult[2]);
                        }
                        var $innerContent=$('.inner-content',$(this));
                        var innerContentMargin=Math.floor($innerContent.height()/2);
                        $innerContent.css('margin-top', '-'+innerContentMargin+'px');
                        $(this).parent().removeClass('plain-text');
                    }
                    else{
                        var postUrl=$(this).find('.post-title a').attr('href');
                        $excerpt.prepend('<div class="post-photo"></div>');
                        var $postPhoto=$('.post-photo',$excerpt);
                        var maskStr='<div class="mask">\
                                        <a class="preview magnific-popup" href="'+$postImg.attr('src')+'">\
                                            <i class="fa fa-search"></i>\
                                        </a>\
                                        <a class="detail" href="'+postUrl+'">\
                                            <i class="fa fa-link"></i>\
                                        </a>\
                                    </div>';
                        $postPhoto.append(maskStr);
                        $postImg.prependTo($postPhoto);
                        $('.mask .magnific-popup').magnificPopup({
                            type: 'image',
                            closeOnContentClick: true,
                            closeBtnInside: false,
                            fixedContentPos: true,
                            mainClass: 'mfp-no-margins mfp-with-zoom', // class to remove default margin from left and right side
                            image: {
                                verticalFit: true
                            }
                        });
                    }
                }
                if($(this).find(".post-excerpt").has('iframe[src^="//www.youtube.com"]').length){
                    var $excerpt=$('.post-excerpt',$(this));
                    var $youtube=$excerpt.find('iframe[src^="//www.youtube.com"]');
                    if(isGallery){
                        var regExp=/youtube(-nocookie)?\.com\/(embed|v)\/([\w_-]+)/;
                        var youtubeId ='';
                        var regResult= $youtube.attr('src').match(regExp);
                        if(regResult.length)
                            youtubeId=regResult[3];
                        if(youtubeId!=''){
                            var youtubeThumnailUrl='http://i3.ytimg.com/vi/'+youtubeId+'/0.jpg';
                            $excerpt.css('background-image','url("'+youtubeThumnailUrl+'")');
                            $(this).parent().removeClass('plain-text');
                        }
                    }
                    else{
                        $excerpt.prepend('<div class="post-video"></div>');
                        var $postYoutube=$('.post-video',$excerpt);
                        $youtube.prependTo($postYoutube);
                    }
                }
                if($(this).find(".post-excerpt").has('iframe[src^="https://w.soundcloud.com"]').length){
                    var $excerpt=$('.post-excerpt',$(this));
                    var $soundcloud=$excerpt.find("iframe:first-child");
                    if(isGallery){
                        var regExp=/soundcloud.com\/tracks\/(\d+)/;
                        var soundcloudId='';
                        var regResult= $soundcloud.attr('src').match(regExp);
                        if(regResult.length)
                            soundcloudId=regResult[1];
                        if(soundcloudId!=''){
                            var soundcloudApiUrl='http://api.soundcloud.com/tracks/'+soundcloudId+'.json?client_id=425fc6ee65a14efbb9b83b1c49a87ccb';
                            $.ajax({
                                type: 'GET',
                                url: soundcloudApiUrl,
                                dataType: "json",
                                success: function(result) {
                                    var artwork_url=result.artwork_url.replace('-large','-t500x500');
                                    $excerpt.css('background-image','url("'+artwork_url+'")');
                                    $excerpt.parent().parent().removeClass('plain-text');
                                }
                            });
                        }
                    }
                    else{
                        $excerpt.prepend('<div class="post-audio"></div>');
                        var $postSoundcloud=$('.post-audio',$excerpt);
                        $soundcloud.prependTo($postSoundcloud);
                    }
                }
                if($(this).find(".post-excerpt").has('iframe[src^="//player.vimeo.com"]').length){
                    var $excerpt=$('.post-excerpt',$(this));
                    var $vimeo=$excerpt.find('iframe[src^="//player.vimeo.com"]');
                    if(isGallery){
                        // get vimeo thumnail
                        var regExp = /video\/(\d+)/;
                        var vimeoId ='';
                        var regResult= $vimeo.attr('src').match(regExp);
                        if(regResult.length)
                            vimeoId=regResult[1];
                        if(vimeoId!=''){
                            var vimeoUrl='http://vimeo.com/api/v2/video/'+vimeoId+'.json';
                            $.ajax({
                                type: 'GET',
                                url: vimeoUrl,
                                dataType: "json",
                                success: function(result) {
                                    if(result.length){
                                        $excerpt.css('background-image','url("'+result[0].thumbnail_large+'")');
                                        $excerpt.parent().parent().removeClass('plain-text');
                                    }
                                }
                            });
                        }
                    }
                    else{
                        $excerpt.prepend('<div class="post-video"></div>');
                        var $postVimeo=$('.post-video',$excerpt);
                        $vimeo.prependTo($postVimeo);
                    }
                }
                else{
                    if(isGallery){
                        var $innerContent=$('.inner-content',$(this));
                        var innerContentMargin=Math.floor($innerContent.height()/2);
                        $innerContent.css('margin-top', '-'+innerContentMargin+'px');
                    }
                }
                $(this).addClass('formated');
            });
            if($('body').is('.enable-filter')){
                ghostApp.filterSetup();
            }
            ghostApp.updateCommentCount();
        }
    },
    getMaxPagination:function(){
        if($('.page-number').length){
            var numberPattern = /\d+/g;
            var pageNumberStr=$('.page-number').html();
            var result=pageNumberStr.match(numberPattern);
            if(result!=null && result.length>1){
                return result[1];
            }
            else{
                return 1;
            }
        }
    },
    getRecentPosts:function(){
        if($('.recent-post').length){
            $('.recent-post').each(function(){
                var $this=$(this);
                var showPubDate=false;
                var showDesc=false;
                var descCharacterLimit=-1;
                var size=-1;
                var type='static';
                var slideMode='horizontal';
                var slideSpeed=500;
                var slidePager=false;
                var isTicker=false;
                var monthName=new Array();
                monthName[0]="Jan";
                monthName[1]="Feb";
                monthName[2]="Mar";
                monthName[3]="Apr";
                monthName[4]="May";
                monthName[5]="June";
                monthName[6]="July";
                monthName[7]="Aug";
                monthName[8]="Sept";
                monthName[9]="Oct";
                monthName[10]="Nov";
                monthName[11]="Dec";
                if($this.data('pubdate'))
                    showPubDate=$this.data('pubdate');
                if($this.data('desc')){
                    showDesc=$this.data('desc');
                    if($this.data('character-limit'))
                        descCharacterLimit=$this.data('character-limit');
                }
                if($this.data('size'))
                    size=$this.data('size');
                if($this.data('type'))
                    type=$this.data('type');
                if(type==='scroll'){
                    if($this.data('mode'))
                        slideMode=$this.data('mode');
                    if($this.data('speed'))
                        slideSpeed=$this.data('speed');
                    if($this.data('pager'))
                        slidePager=$this.data('pager');
                    if($this.data('ticker'))
                        isTicker=$this.data('ticker');
                }
                $.ajax({
                    type: 'GET',
                    url: '/rss/',
                    dataType: "xml",
                    success: function(xml) {
                        if($(xml).length){
                            var htmlStr='';
                            var date;
                            var count=0;
                            $('item', xml).each( function() {
                                if(size>0 && count < size){
                                    htmlStr+='<li class="clearfix">';
                                    if(showPubDate){
                                        date = new Date($(this).find('pubDate').eq(0).text());
                                        htmlStr += '<span class="itemDate">\
                                                        <span class="date">'+date.getDate()+'</span>\
                                                        <span class="month">'+monthName[date.getMonth()]+'</span>\
                                                    </span>';
                                    }
                                    htmlStr+='<span class="itemContent">';
                                    htmlStr += '<span class="title">\
                                                        <a href="' + $(this).find('link').eq(0).text() + '">\
                                                        ' + $(this).find('title').eq(0).text() + '\
                                                        </a>\
                                                </span>';
                                    if (showDesc) {
                                        var desc=$(this).find('description').eq(0).text();
                                        // trip html tag
                                        desc=$(desc).text();
                                        if (descCharacterLimit > 0 && desc.length > descCharacterLimit) {
                                            htmlStr += '<span class="desc">' + desc.substr(0, descCharacterLimit) + ' ...\
                                                            <a href="'+$(this).find('link').eq(0).text()+'">View »</a>\
                                                        </span>';
                                        }
                                        else{
                                            htmlStr += '<span class="desc">' + desc + "</span>";
                                        }
                                    }
                                    htmlStr+='</span>';
                                    htmlStr += '</li>';
                                    count++;
                                }
                                else{
                                    return false;
                                }
                            });
                            if(type==='static')
                                htmlStr='<ul class="feedList static">'+ htmlStr + "</ul>";
                            else{
                                htmlStr='<ul class="feedList bxslider">'+ htmlStr + "</ul>";
                            }
                            $this.append(htmlStr);
                            if(type!=='static'){
                                // Updating on v1.2
                            }
                        }
                    },
                    error: function (xhr, ajaxOptions, thrownError) {
                        console.log(thrownError);
                    }
                });
            });
        }
    },
    getFlickr:function(){
        if($('.flickr-feed').length){
            $('.flickr-feed').each(function() {
                var $this=$(this);
                if(flickr_id=='' || flickr_id=='YOUR_FLICKR_ID_HERE'){
                    $(this).html('<li><strong>Please change Flickr user id before use this widget</strong></li>');
                }
                else{
                    var feedTemplate='<li><a href="{{image_b}}"><img src="{{image_m}}" alt="{{title}}" /></a></li>';
                    var size=15;
                    if($(this).data('size'))
                        size=$(this).data('size');
                    var isPopupPreview=false;
                    if($(this).data('popup-preview'))
                        isPopupPreview=$(this).data('popup-preview');
                    $(this).jflickrfeed({
                        limit: size,
                        qstrings: {
                            id: flickr_id
                        },
                        itemTemplate: feedTemplate
                    }, function(data) {
                        if(isPopupPreview){
                            $this.magnificPopup({
                                delegate: 'a',
                                type: 'image',
                                tLoading: 'Loading image #%curr%...',
                                closeOnContentClick: true,
                                closeBtnInside: false,
                                fixedContentPos: true,
                                mainClass: 'mfp-no-margins mfp-with-zoom', // class to remove default margin from left and right side
                                image: {
                                    verticalFit: true,
                                    tError: '<a href="%url%">The image #%curr%</a> could not be loaded.'
                                },
                                gallery: {
                                    enabled: true,
                                    navigateByImgClick: true,
                                    preload: [0,1] // Will preload 0 - before current, and 1 after the current image
                                }
                            });
                        }
                    });
                }
            });
        }
    },
    getInstagram: function(){
        if($('.instagram-feed').length){
            if(instagram_accessToken!='' || instagram_accessToken!='your-instagram-access-token'){
                /* Initalize Instagram Feed */
                $.fn.spectragram.accessData = {
                    accessToken: instagram_accessToken,
                    clientID: instagram_clientID
                };
            }
            $('.instagram-feed').each(function(){
                if(instagram_accessToken=='' || instagram_accessToken=='your-instagram-access-token'){
                    $(this).html('<li><strong>Please change instagram api access info before use this widget</strong></li>');
                }
                else{
                    var display=15;
                    var wrapEachWithStr='<li></li>';
                    if($(this).data('display'))
                        display=$(this).data('display');
                    $(this).spectragram('getRecentTagged',{
                        query: 'converse',
                        max: display
                    });
                }
            });
        }
    },
    getDribbble: function(){
        if($('.dribbble-feed').length){
            $('.dribbble-feed').each(function(){
                var $this=$(this);
                var display=15;
                if($this.data('display'))
                    display=$this.data('display');
                var isPopupPreview=false;
                if($this.data('popup-preview'))
                    isPopupPreview=$this.data('popup-preview');
                $.jribbble.getShotsByList('popular', function (listDetails) {
                    var html = [];
                    $.each(listDetails.shots, function (i, shot) {
                        html.push('<li><a href="' + shot.url + '"><img src="' + shot.image_teaser_url + '" alt="' + shot.title + '"></a></li>');
                    });
                    $this.html(html.join(''));
                    if(isPopupPreview){
                        $this.magnificPopup({
                            delegate: 'a',
                            type: 'image',
                            tLoading: 'Loading image #%curr%...',
                            closeOnContentClick: true,
                            closeBtnInside: false,
                            fixedContentPos: true,
                            mainClass: 'mfp-no-margins mfp-with-zoom', // class to remove default margin from left and right side
                            image: {
                                verticalFit: true,
                                tError: '<a href="%url%">The image #%curr%</a> could not be loaded.'
                            },
                            gallery: {
                                enabled: true,
                                navigateByImgClick: true,
                                preload: [0,1] // Will preload 0 - before current, and 1 after the current image
                            }
                        });
                    }
                }, {page: 1, per_page: display});
            });
        }
    },
    isotopeSetup:function(){
        if($('#isotope-content').length){
            var $container = $('#isotope-content');
            $('.post').imagesLoaded(function( instance ) {
                ghostApp.formatWidth();
                if($('body').is('.gallery')){
                    var columns=3;
                    var contentWidth = $container.width();
                    if (contentWidth<=480){
                        columns=1;
                    }
                    else if (contentWidth<=640){
                        columns=2;
                    }
                    else if (contentWidth<=800){
                        columns=3;
                    }
                    else if(contentWidth<=1140){
                        columns=4;
                    }
                    else if(contentWidth<=1380){
                        columns=5;
                    }
                    else if(contentWidth<=1620){
                        columns=6;
                    }
                    else if(contentWidth>1620){
                        columns=7;
                    }
                    $container.isotope({
                        itemSelector : '.item-list',
                        layoutMode: 'perfectMasonry',
                        resizable: false,
                        animationOptions: {
                            duration: 400,
                            easing: 'swing',
                            queue: false
                        },
                        perfectMasonry: {
                            layout: 'vertical',
                            liquid: true,
                            columnWidth: 200,   // Force columns to 200px wide
                            rowHeight: 200,      // Force rows to 200px high
                            cols: columns
                        }
                    });
                }
                else{
                    if($('body').is('.enable-filter')){
                        $container.isotope({
                            itemSelector : '.item-list',
                            filter:'*',
                            resizable: false,
                            animationOptions: {
                                duration: 400,
                                easing: 'swing',
                                queue: false
                            },
                            getSortData: {
                                date: function ($elem) {
                                    return Date.parse($elem.data('publish-date'));
                                },
                                title:function($elem){
                                    return $elem.find('.post-title a').text().replace(' ','');
                                }
                            }
                        });
                    }
                    else{
                        $container.isotope({
                            itemSelector : '.item-list',
                            resizable: false,
                            animationOptions: {
                                duration: 400,
                                easing: 'swing',
                                queue: false
                            }
                        });
                    }
                }
                $('#preloader').fadeOut('fast', function () {
                    $(this).remove();
                    if($('body').is('.enable-filter')){
                        $('.filter-control').fadeIn('fast', function() {
                            var marginTop=$('.filter-control').height()+40;
                            if($('body').is('.gallery'))
                                marginTop=marginTop-13;
                            $('#isotope-content').css('margin-top', marginTop+'px');
                            $('.filter-control').animate({'top': '0px'},
                                400,
                                function() {
                                    $('#isotope-content').fadeIn('slow').css({
                                        'visibility':'visible',
                                    });
                                ghostApp.reloadIsotope();
                            });
                        });
                    }
                    else{
                        $('#isotope-content').fadeIn('slow').css({
                            'visibility':'visible',
                        });
                        ghostApp.reloadIsotope();
                    }
                });
            });
        }
    },
    infiniteScrollSetup:function(){
        if($('#isotope-content').length){
            var $container = $('#isotope-content');
            $container.infinitescroll({
                    navSelector     : '.pagination',    // selector for the paged navigation
                    nextSelector    : '.pagination a',  // selector for the NEXT link (to page 2)
                    itemSelector    : '.post',     // selector for all items you'll retrieve
                    maxPage           : ghostApp.getMaxPagination(),
                    loading: {
                        finishedMsg: 'No more post to load.',
                            img: '/assets/img/loading.gif'
                    }
                },
                // call Isotope as a callback
                function( newElements ) {
                    $container.isotope('insert', $(newElements),function(){
                        ghostApp.responsiveImages();
                        ghostApp.reformatPost();
                        $(".post .wrap").fitVids();
                        ghostApp.reloadIsotope();
                    });
                }
            );
        }
    },
    reloadIsotope:function(){
        if($('#isotope-content').length){
            ghostApp.formatWidth();
            if($('body').is('.gallery')){
                var columns=3;
                var contentWidth = $('.body-content').width();
                if (contentWidth<=480){
                    columns=1;
                }
                else if (contentWidth<=640){
                    columns=2;
                }
                else if (contentWidth<=800){
                    columns=3;
                }
                else if(contentWidth<=1140){
                    columns=4;
                }
                else if(contentWidth<=1380){
                    columns=5;
                }
                else if(contentWidth<=1620){
                    columns=6;
                }
                else if(contentWidth>1620){
                    columns=7;
                }
                ghostApp.centerPostDesc();
                $('#isotope-content').isotope({
                        perfectMasonry: {
                            cols: columns
                        }
                });
            }
            else{
                $('#isotope-content').isotope();
            }
        }
    },
    fitVids:function(){
        $(".post .wrap").fitVids();
        $(".post").fitVids();
    },
    updateCommentCount:function(){
        if($('.disqus-comment-count').length){
            var urlArray = [];
            $('.disqus-comment-count').each(function () {
              var url = $(this).attr('data-disqus-url');
              urlArray.push('link:' + url);
            });
            try{
                $.get("http://disqus.com/api/3.0/threads/set.jsonp",{
                    api_key: disqusPublicKey,
                    forum : disqusShortname,
                    thread : urlArray},
                    function (result) {
                        for (var i in result.response) {
                            var countText = " comments";
                            var count = result.response[i].posts;
                            if (count <= 1){
                                countText = " comment";
                            }
                            $('span[data-disqus-url="' + result.response[i].link + '"]').html(count + countText);
                        }
                    },
                "jsonp");
            }
            catch(err){
                console.log(err);
            }
        }
        else if(window.FB){
            FB.XFBML.parse(document.body);
        }
        else if($('.gplus-comment-count').length){
            var count=1;
            $('.gplus-comment-count').each(function(){
                $(this).attr('id', 'commentscounter'+count);
                gapi.commentcount.render('commentscounter'+count, {
                    href: $(this).data('href'),
                    color: $(this).data('color')
                });
                count++;
            });
        }
    },
    contentPopup:function(){
        if($('.main-nav').length){
            $('.main-nav li a').each(function(index, el) {
                if($(this).is('.content-popup')){
                    $(this).magnificPopup({
                        type: 'inline',
                        fixedContentPos: false,
                        fixedBgPos: true,
                        overflowY: 'auto',
                        closeBtnInside: true,
                        preloader: false,
                        midClick: true,
                        removalDelay: 300,
                        mainClass: 'my-mfp-zoom-in'
                    });
                }
                else if($(this).is('.content-popup-ajax') && $(this).attr('href')!=''){
                    $(this).magnificPopup({
                        type: 'ajax',
                        fixedContentPos: false,
                        fixedBgPos: true,
                        overflowY: 'auto',
                        closeBtnInside: true,
                        preloader: false,
                        midClick: true,
                        removalDelay: 300,
                        mainClass: 'my-mfp-zoom-in',
                        ajax: {
                            settings: null,
                            cursor: 'mfp-ajax-cur',
                            tError: '<a href="%url%">The content</a> could not be loaded.'
                        },
                        callbacks: {
                            parseAjax: function(mfpResponse) {
                                var popupContentStr='<div class="popup-container">';
                                popupContentStr+=$(mfpResponse.data).find('.post-title').html();
                                popupContentStr+=$(mfpResponse.data).find('.post-content').html();
                                popupContentStr+='</div>';
                                mfpResponse.data = popupContentStr;
                            }
                        }
                    });
                }
            });
        }
        else if($('.bottom-nav').length){
            $('.bottom-nav li a').each(function(index, el) {
                if($(this).is('.content-popup')){
                    $(this).magnificPopup({
                        type: 'inline',
                        fixedContentPos: false,
                        fixedBgPos: true,
                        overflowY: 'auto',
                        closeBtnInside: true,
                        preloader: false,
                        midClick: true,
                        removalDelay: 300,
                        mainClass: 'my-mfp-zoom-in'
                    });
                }
                else if($(this).is('.content-popup-ajax') && $(this).attr('href')!=''){
                    $(this).magnificPopup({
                        type: 'ajax',
                        fixedContentPos: false,
                        fixedBgPos: true,
                        overflowY: 'auto',
                        closeBtnInside: true,
                        preloader: false,
                        midClick: true,
                        removalDelay: 300,
                        mainClass: 'my-mfp-zoom-in',
                        ajax: {
                            settings: null,
                            cursor: 'mfp-ajax-cur',
                            tError: '<a href="%url%">The content</a> could not be loaded.'
                        },
                        callbacks: {
                            parseAjax: function(mfpResponse) {
                                var popupContentStr='<div class="popup-container">';
                                popupContentStr+=$(mfpResponse.data).find('.post-title').html();
                                popupContentStr+=$(mfpResponse.data).find('.post-content').html();
                                popupContentStr+='</div>';
                                mfpResponse.data = popupContentStr;
                            }
                        }
                    });
                }
            });
        }
    },
    initFilter:function(){
        if($('body').is('.enable-filter') && $('.filter-control').length){
            var noSelectedText="None Selected";
            if($(window).width()<=480){
                noSelectedText="None";
            }
            $('#filter').multiselect({
                numberDisplayed: 1,
                buttonText: function(options, select) {
                    if (options.length == 0) {
                      return noSelectedText+' <b class="caret"></b>';
                    }
                    else {
                        return options.length + ' selected <b class="caret"></b>';
                    }
                },
                onChange: function(option, checked) {
                    var selected='';
                    $('#filter option:selected').each(function() {
                        if(selected!='')
                            selected+=','+$(this).val();
                        else
                            selected+=$(this).val();
                    });
                    if(selected=='' || selected.indexOf('*')>=0){
                        selected='*';
                    }
                    $('#isotope-content').isotope({filter:selected});
                }
            });
            $('#sort').multiselect({
                onChange: function(option, checked) {
                    var sortAsc=false;
                    if($('.sort-option').data('sort')=='asc')
                        sortAsc=true;
                    var sortByStr=option.val();
                    console.log('sort by '+sortByStr+' order by '+$('.sort-option').data('sort') +' ('+sortAsc+')');
                    $('#isotope-content').isotope({
                        sortBy : sortByStr,
                        sortAscending : sortAsc
                    });
                }
            });
            $('.sort-option').click(function(event) {
                var sortAsc=false;
                var sortAscStr='';
                if($(this).data('sort')=='asc'){
                    sortAsc=false;
                    sortAscStr='desc';
                    $(this).data('sort','desc');
                    $(this).html('<i class="fa fa-long-arrow-down"></i>');
                }
                else{
                    sortAsc=true;
                    sortAscStr='asc';
                    $(this).data('sort','asc');
                    $(this).html('<i class="fa fa-long-arrow-up"></i>');
                }
                var sortByStr=$('#sort').val();
                console.log('sort by '+sortByStr+' order by '+$(this).data('sort') +' ('+sortAsc+')');
                $('#isotope-content').isotope({
                    sortBy : sortByStr,
                    sortAscending : sortAsc
                });
            });
        }
    },
    filterSetup:function(){
        if($('.filter-control').length){
            var $filterControl=$('#filter');
            $filterControl.html('');
            var lastActiveTag='*';
            var filters=[];
            var i=0;
            var filterItemStr='';
            var activeClassStr='';
            $('.post').each(function(){
                var tagList='';
                if($(this).data('tags')){
                    tagList=$(this).data('tags');
                }
                if(tagList.length>0){
                    var tags=tagList.split(',');
                    for (i = 0; i < tags.length; i++) {
                        if(tags[i]!=''){
                            var tagOj= {"tagName": tags[i],"total": 1};
                            var hasOld=false;
                            for(var j=0; j<filters.length; j++){
                                if(filters[j].tagName==tags[i]){
                                    tagOj.total=filters[j].total+1;
                                    filters[j]=tagOj;
                                    hasOld=true;
                                }
                            }
                            if(!hasOld)
                                filters.push(tagOj);
                        }
                    }
                }
            });
            for (i = 0; i < filters.length; i++) {
                activeClassStr='';
                var curFilterStr='.tag-'+filters[i].tagName.toLowerCase().replace(' ','-');
                if(lastActiveTag!=='*' && rememberLastActiveTag && lastActiveTag===curFilterStr){
                    activeClassStr='selected';
                }
                filterItemStr=' <option value="'+curFilterStr+'" '+activeClassStr+'>'+filters[i].tagName+'</option>';
                $filterControl.append(filterItemStr);
            }
            $filterControl.multiselect('rebuild');
        }
    },
    mainMenuEvents:function(){
        if($('.main-nav').length){
            var currentUrl=window.location.href;
            var $currentMenu=$('.main-nav a[href="'+currentUrl+'"]');
            if($currentMenu.length){
                $('.main-nav li.active').removeClass('active');
                $currentMenu.parent().addClass('active');
                var $parentMenu=$currentMenu.parent().closest('.has-sub');
                if($parentMenu.length){
                    $parentMenu.addClass('active');
                }
            }
            $('.main-nav>li.has-sub>a').click(function(e) {
                e.preventDefault();
                var targetDropdown = $(this).next('.dropdown-menu');
                var targetParent = $(targetDropdown).closest('.has-sub');
                if($(targetDropdown).is(':hidden')) {
                    $.when($(targetDropdown).slideToggle(200)).done(function() {
                        $(targetDropdown).closest('.has-sub').addClass('active');
                    });
                } else {
                    $.when($(targetDropdown).slideUp(200)).done(function() {
                        $(targetDropdown).closest('.has-sub').removeClass('active');
                    });
                }
            });
        }
    },
    syntaxHighlighter:function(){
        SyntaxHighlighter.autoloader.apply(null, path(
            'applescript            @shBrushAppleScript.js',
            'actionscript3 as3      @shBrushAS3.js',
            'bash shell             @shBrushBash.js',
            'coldfusion cf          @shBrushColdFusion.js',
            'cpp c                  @shBrushCpp.js',
            'c# c-sharp csharp      @shBrushCSharp.js',
            'css                    @shBrushCss.js',
            'delphi pascal          @shBrushDelphi.js',
            'diff patch pas         @shBrushDiff.js',
            'erl erlang             @shBrushErlang.js',
            'groovy                 @shBrushGroovy.js',
            'java                   @shBrushJava.js',
            'jfx javafx             @shBrushJavaFX.js',
            'js jscript javascript  @shBrushJScript.js',
            'perl pl                @shBrushPerl.js',
            'php                    @shBrushPhp.js',
            'text plain             @shBrushPlain.js',
            'py python              @shBrushPython.js',
            'powershell ps posh     @shBrushPowerShell.js',
            'ruby rails ror rb      @shBrushRuby.js',
            'sass scss              @shBrushSass.js',
            'scala                  @shBrushScala.js',
            'sql                    @shBrushSql.js',
            'vb vbnet               @shBrushVb.js',
            'xml xhtml xslt html    @shBrushXml.js'
        ));
        SyntaxHighlighter.all();
    },
    mailchimpSetup:function(){
        // Submit event
        $("#mc-form input").not("[type=submit]").jqBootstrapValidation({
            submitSuccess: function ($form, event) {
                var url=$form.attr('action');
                event.preventDefault();
                if(url=='' || url=='YOUR_WEB_FORM_URL_HERE')
                {
                    alert('Please config your mailchimp form url for this widget');
                    return false;
                }
                else{
                    url=url.replace('/post?', '/post-json?').concat('&c=?');
                    var data = {};
                    var dataArray = $form.serializeArray();
                    $.each(dataArray, function (index, item) {
                        data[item.name] = item.value;
                    });
                    $.ajax({
                        url: url,
                        data: data,
                        success: function(resp){
                            if (resp.result === 'success') {
                                alert("Got it, you've been added to our newsletter. Thanks for subscribe!");
                            }
                            else{
                                alert(resp.result);
                            }
                        },
                        dataType: 'jsonp',
                        error: function (resp, text) {
                            console.log('mailchimp ajax submit error: ' + text);
                        }
                    });
                    return false;
                }
            }
        });
    },
    postAnimation:function(){
        if($('body').is('.post-template')){
            var headerTop=0;
            var postMargin=0;
            if($(window).width()<=979){
                headerTop=69;
                postMargin=$('.post-header').height()+69;
            }
            $('.post-wrap').imagesLoaded(function( instance ) {
                $('#preloader').stop().fadeOut('fast', function () {
                    $(this).remove();
                    $('.post-header').animate({
                        top: headerTop+'px'
                    },50, function() {
                        $('.comment-box').show();
                        $('.comment-box').niceScroll();
                        $('.post-wrap').fadeIn(1000);
                        $('.body-content').css('margin-top', postMargin+'px');
                    });
                });
            });
        }
    },
    gmapInitialize:function(){
        if($('.gmap').length){
            var mainColor=ghostApp.hexColor($('.post-header').css('backgroundColor'));
            var myLatlng = new google.maps.LatLng(your_latitude,your_longitude);
            var mapOptions = {
                zoom: 13,
                scrollwheel: false,
                center: myLatlng,
                styles: [{"stylers":[{"hue": mainColor, "lightness" : 100}]}]
            }
            var map = new google.maps.Map(document.getElementById('gmap'), mapOptions);
            var marker = new google.maps.Marker({
                position: myLatlng,
                map: map,
                title: 'Form'
            });
        }
    },
    hexColor:function(colorval) {
        var parts = colorval.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
        delete(parts[0]);
        for (var i = 1; i <= 3; ++i) {
            parts[i] = parseInt(parts[i]).toString(16);
            if (parts[i].length == 1) parts[i] = '0' + parts[i];
        }
        return '#' + parts.join('');
    },
    searchFormHandle:function(){
        if($('.search-box').length){
            $('#search-keyword').keyup(function () {
                if($('#search-keyword').val().length>3){
                    $.getJSON(
                        'http://tapirgo.com/api/1/search.json?token=' + tapir_token + '&query=' + $('#search-keyword').val() + '&callback=?', function(data){
                            console.log(data);
                            var resultStr='<span class="no-result">No results were found for your search.</span>';
                            if(data.length){
                                resultStr='<ul>';
                                $.each(data, function(key, val) {
                                    var itemSummary=$(val.summary).text();
                                    if(itemSummary.length>search_haracter_limit)
                                        itemSummary=itemSummary.substr(0, search_haracter_limit) +' ...';
                                    resultStr+='<li>\
                                                <a href="'+ val.link +'">\
                                                    <span class="title">'+val.title+'</span>\
                                                    <span class="desc">'+itemSummary+'</span>\
                                                </a>\
                                            </li>';
                                });
                                resultStr+='</ul>';
                            }
                            $('#search-result').html(resultStr).fadeIn('fast');
                        }
                    );
                }
                else{
                    $('#search-result').html('');
                }
            });
        }
    },
    scrollEvent:function(){
        $(window).scroll(function() {
            "use strict";
            var curPos = $(window).scrollTop();
            if($('body').is('.post-template')){
                if (curPos >= $('.post-header').height()) {
                    $('.post-header').addClass('fixed-top');
                } else {
                    $('.post-header').removeClass('fixed-top');
                }
            }
            else{
                if($(window).width()<=979){
                    if(curPos>=$('#main-sidebar').height()){
                        $('.filter-control').addClass('fixed');
                    }
                    else{
                        $('.filter-control').removeClass('fixed');
                    }
                }
            }
        });
    },
    uiInit:function(){
        $('html, document, body').scrollTop(0);
        $('#isotope-content').hide();
        // Center Intro
        if($('body').is('.simple-style') || $(window).width()<=979){
            var introHeight=$('#main-sidebar .intro').height();
            $('#main-sidebar .intro').css({
                'margin-top': '-'+Math.floor(introHeight/2)+'px'
            }).fadeIn('slow', function() {});
        }
    },
    misc:function(){
        $('#mobile-menu').click(function(event) {
            if(!$('body').is('.open-menu')){
                $('body').addClass('open-menu');
                if($('.post-header').length){
                    var headerTop=$('.post-header').position().top;
                    var menuHeight=$('.main-nav').height();
                    headerTop+=menuHeight;
                    $('.post-header').css('top', headerTop+'px');
                }
            }
            else{
                $('body').removeClass('open-menu');
                if($('.post-header').length){
                    if( $('.post-header').is('.fixed-top')){
                        $('.post-header').css('top', '1px');
                    }
                    else{
                        $('.post-header').css('top', '69px');
                    }
                }
            }
        });
        $('#more-info').click(function(event) {
            if($('body').is('.open-more-info')){
                $(this).find('i.fa').removeClass('fa-minus');
                $(this).find('i.fa').addClass('fa-plus');
                $('body').removeClass('open-more-info');
                $('.nicescroll-rails').hide();
            }
            else{
                $(this).find('i.fa').removeClass('fa-plus');
                $(this).find('i.fa').addClass('fa-minus');
                $('body').addClass('open-more-info');
                if($('.nicescroll-rails').length)
                    $('.nicescroll-rails').show();
            }
        });
        // BackToTop Button click event
        $('.totop').click(function () {
            $("html, body").animate({scrollTop: 0}, 800);
            return false;
        });
        $('#turncomment').click(function(event) {
            if($('body').is('.opencomment')){
                $('body').removeClass('opencomment');
            }
            else{
                $('body').addClass('opencomment');
            }
            if($(window).width()<=979){
                $('html, body').stop().animate({scrollTop: $('.comment-box').offset().top-$('.post-header').height()}, 1000,'easeInOutExpo');
            }
        });
        $('#main-content').click(function(event) {
            if($('body').is('.open-more-info')){
                $('#more-info').trigger('click');
            }
        });
        $('#more-info-sidebar').niceScroll();
        if($('.share-box').length){
            $('#btn-share').click(function(event) {
                event.preventDefault();
                if($('.share-box-content').is(':hidden')){
                    $('body').addClass('open-share');
                    $('.share-box-content').fadeIn('fast');
                }
                else{
                    $('.share-box-content').fadeOut('fast');
                    $('body').removeClass('open-share');
                }
            });
        }
        if($('.search-box').length){
            $('#btn-search').click(function(event) {
                event.preventDefault();
                if($('.search-form').is(':hidden')){
                    $('body').addClass('open-search');
                    $('.search-form').fadeIn('fast');
                }
                else{
                    $('.search-form').fadeOut('fast');
                    $('#search-keyword').val('');
                    $('#search-result').html('');
                    $('body').removeClass('open-search');
                }
            });
        }
        if($('.gmap').length){
            ghostApp.gmapInitialize();
            google.maps.event.addDomListener(window, 'load', ghostApp.gmapInitialize);
            google.maps.event.addDomListener(window, 'resize', ghostApp.gmapInitialize);
        }
    },
    init: function () {
        ghostApp.uiInit();
        ghostApp.initFilter();
        ghostApp.responsiveImages();
        ghostApp.reformatPost();
        ghostApp.fitVids();
        ghostApp.isotopeSetup();
        ghostApp.infiniteScrollSetup();
        ghostApp.getRecentPosts();
        ghostApp.getFlickr();
        ghostApp.getInstagram();
        ghostApp.getDribbble();
        ghostApp.syntaxHighlighter();
        ghostApp.contentPopup();
        ghostApp.mailchimpSetup();
        ghostApp.postAnimation();
        ghostApp.mainMenuEvents();
        ghostApp.searchFormHandle();
        ghostApp.scrollEvent();
        ghostApp.misc();
    }
};

/*================================================================*/
/*  2. Initialing
/*================================================================*/
$(document).ready(function() {
    ghostApp.init();
});
$(window).on("debouncedresize", function( event ) {
    ghostApp.reloadIsotope();
});
$(window).load(function(){
    ghostApp.reloadIsotope();
});
$(window).resize(function () {
    "use strict";
    ghostApp.reloadIsotope();
});