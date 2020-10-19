using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json;

namespace GoogleLibrary
{
    public class GoogleImages
    {
        private string _searchParams;
        private const string GOOGLE_IMAGE_SEARCH_API_URI =
            @"https://www.googleapis.com/customsearch/v1?key=AIzaSyBTwASr75UbGxxT9qSJJrOkVNxk3_1bOvI&cx=faa9aa8097b8b7a8d&q=";

        public GoogleImages()
        {
        }

        public void SearchGoogleImages(List<string> searchParams)
        {
            StringBuilder sb = new StringBuilder();
            string searchString = "";
            for (int i = 0; i < searchParams.Count; ++i)
            {
                sb.Append(searchParams[i]);
                if (i != searchParams.Count - 1)
                {
                    // if your not at the last index append a "+" to have correct google search format for multiple params
                    sb.Append("+");
                }
            }
            _searchParams = sb.ToString();
            GData = InitlializeGoogleImagesFromWebApi();
        }

        public GoogleData GData { get; set; }

        private GoogleData InitlializeGoogleImagesFromWebApi()
        {
            using (WebClient client = new WebClient())
            {
                GoogleData gData = new GoogleData();
                try
                {
                    string jsonData = client.DownloadString($"{GOOGLE_IMAGE_SEARCH_API_URI}{_searchParams}");
                    gData = JsonConvert.DeserializeObject<GoogleData>(jsonData);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e);
                }

                return gData;
            }
        }
    }
}
