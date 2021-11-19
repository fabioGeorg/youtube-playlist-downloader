using System;

namespace YoutubePlaylistDownloader
{
    public class VideoData
    {
        public string VideoTitle { get; protected set; }
        public string VideoURL { get; protected set; }

        public VideoData(string videoTitle, string videoUrl)
        {
            VideoTitle = videoTitle;
            VideoURL = videoUrl;
        }

    }
}
