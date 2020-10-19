using System;
using ILibrary;

namespace GoogleLibrary
{
    public class GoogleData
    {
        public string kind { get; set; }
        public Url url { get; set; }
        public Queries queries { get; set; }
        public Context context { get; set; }
        public Searchinformation searchInformation { get; set; }
        public Item[] items { get; set; }
    }

    public class Url
    {
        public string type { get; set; }
        public string template { get; set; }
    }

    public class Queries
    {
        public Request[] request { get; set; }
        public Nextpage[] nextPage { get; set; }
    }

    public class Request
    {
        public string title { get; set; }
        public string totalResults { get; set; }
        public string searchTerms { get; set; }
        public int count { get; set; }
        public int startIndex { get; set; }
        public string inputEncoding { get; set; }
        public string outputEncoding { get; set; }
        public string safe { get; set; }
        public string cx { get; set; }
    }

    public class Nextpage
    {
        public string title { get; set; }
        public string totalResults { get; set; }
        public string searchTerms { get; set; }
        public int count { get; set; }
        public int startIndex { get; set; }
        public string inputEncoding { get; set; }
        public string outputEncoding { get; set; }
        public string safe { get; set; }
        public string cx { get; set; }
    }

    public class Context
    {
        public string title { get; set; }
    }

    public class Searchinformation
    {
        public float searchTime { get; set; }
        public string formattedSearchTime { get; set; }
        public string totalResults { get; set; }
        public string formattedTotalResults { get; set; }
    }

    public class Item
    {
        public string kind { get; set; }
        public string title { get; set; }
        public string htmlTitle { get; set; }
        public string link { get; set; }
        public string displayLink { get; set; }
        public string snippet { get; set; }
        public string htmlSnippet { get; set; }
        public string cacheId { get; set; }
        public string formattedUrl { get; set; }
        public string htmlFormattedUrl { get; set; }
        public Pagemap pagemap { get; set; }
    }

    public class Pagemap
    {
        public Cse_Thumbnail[] cse_thumbnail { get; set; }
        public Metatag[] metatags { get; set; }
        public Cse_Image[] cse_image { get; set; }
        public Webpage[] webpage { get; set; }
        public Imageobject[] imageobject { get; set; }
        public Person[] person { get; set; }
        public Videoobject[] videoobject { get; set; }
        public Scraped[] scraped { get; set; }
        public Wpheader[] wpheader { get; set; }
        public Sitenavigationelement[] sitenavigationelement { get; set; }
    }

    public class Cse_Thumbnail
    {
        public string src { get; set; }
        public string width { get; set; }
        public string height { get; set; }
    }

    public class Metatag
    {
        public string ogimage { get; set; }
        public string themecolor { get; set; }
        public string ogtype { get; set; }
        public string twittercard { get; set; }
        public string twittertitle { get; set; }
        public string ogsite_name { get; set; }
        public string twitterurl { get; set; }
        public string ogtitle { get; set; }
        public string ogdescription { get; set; }
        public string fbapp_id { get; set; }
        public string mobilewikiconfigenvironment { get; set; }
        public string twittersite { get; set; }
        public string viewport { get; set; }
        public string twitterdescription { get; set; }
        public string ogurl { get; set; }
        public string title { get; set; }
        public string twitterdomain { get; set; }
        public string twitterappurliphone { get; set; }
        public string twitterappidgoogleplay { get; set; }
        public string ogimagewidth { get; set; }
        public string twitterappurlipad { get; set; }
        public string alandroidpackage { get; set; }
        public string twitterappnamegoogleplay { get; set; }
        public string aliosurl { get; set; }
        public string twitterappidiphone { get; set; }
        public string aliosapp_store_id { get; set; }
        public string twitterimage { get; set; }
        public string twitterplayer { get; set; }
        public string twitterplayerheight { get; set; }
        public string ogvideotype { get; set; }
        public string ogvideoheight { get; set; }
        public string ogvideourl { get; set; }
        public string aliosapp_name { get; set; }
        public string ogimageheight { get; set; }
        public string twitterappidipad { get; set; }
        public string alweburl { get; set; }
        public string ogvideosecure_url { get; set; }
        public string ogvideotag { get; set; }
        public string ogvideowidth { get; set; }
        public string alandroidurl { get; set; }
        public string twitterappurlgoogleplay { get; set; }
        public string twitterappnameipad { get; set; }
        public string twitterplayerwidth { get; set; }
        public string alandroidapp_name { get; set; }
        public string twitterappnameiphone { get; set; }
        public string referrer { get; set; }
        public DateTime articlepublished_time { get; set; }
        public string articlesection { get; set; }
        public string fbpages { get; set; }
        public DateTime articlemodified_time { get; set; }
        public string oglocale { get; set; }
        
    }

    public class Cse_Image : IData
    {
        public string src { get; set; }
        public string URL { get => src; set => src = value; }
    }

    public class Webpage
    {
        public string image { get; set; }
        public string name { get; set; }
    }

    public class Imageobject
    {
        public string width { get; set; }
        public string url { get; set; }
        public string height { get; set; }
    }

    public class Person
    {
        public string name { get; set; }
        public string url { get; set; }
    }

    public class Videoobject
    {
        public string embedurl { get; set; }
        public string playertype { get; set; }
        public string isfamilyfriendly { get; set; }
        public string uploaddate { get; set; }
        public string description { get; set; }
        public string videoid { get; set; }
        public string url { get; set; }
        public string duration { get; set; }
        public string unlisted { get; set; }
        public string name { get; set; }
        public string paid { get; set; }
        public string width { get; set; }
        public string regionsallowed { get; set; }
        public string genre { get; set; }
        public string interactioncount { get; set; }
        public string channelid { get; set; }
        public string datepublished { get; set; }
        public string thumbnailurl { get; set; }
        public string height { get; set; }
    }

    public class Scraped
    {
        public string image_link { get; set; }
    }

    public class Wpheader
    {
        public string headline { get; set; }
    }

    public class Sitenavigationelement
    {
        public string name { get; set; }
        public string url { get; set; }
    }

}
