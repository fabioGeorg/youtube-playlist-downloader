using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using Google.Apis.YouTube.v3;
using Google.Apis.Services;


namespace YoutubePlaylistDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: YoutubePlaylistDownloader PlaylistID FileName (without extension)");
                return;
            }
            List<VideoData> m_PlaylistTitles = new List<VideoData>();
            int i = 0;            

            try
            {
                //TODO: Fix System.ReferenceException
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = config.AppSettings.Settings;
                string apiKey = settings["apiKey"].Value;
                using (var svcYoutube = new YouTubeService(new BaseClientService.Initializer()
                {
                    ApiKey = apiKey,
                    ApplicationName = "YoutubePlaylistDownloader"
                }))
                {
                    var nextPageToken = "";
                    while (nextPageToken != null)
                    {
                        var playlist = svcYoutube.PlaylistItems.List("snippet");
                        playlist.PlaylistId = args[0];
                        playlist.MaxResults = 50;
                        playlist.PageToken = nextPageToken;

                        var resp = playlist.Execute();
                        foreach(var item in resp.Items)
                        {
                            string videoTitle = item.Snippet.Title;
                            string videoUrl = $"https://www.youtube.com/watch?v={item.Snippet.ResourceId.VideoId}";
                            VideoData tmp = new VideoData(videoTitle, videoUrl);
                            m_PlaylistTitles.Add(tmp);
                            Console.WriteLine($"Done: {++i} - {videoTitle}");
                        }                        
                        nextPageToken = resp.NextPageToken;
                    }

                    //var excelApp = new Application();                    
                    //excelApp.Visible = true;
                    //excelApp.Workbooks.Add();
                    //_Worksheet worksheet = (Worksheet)excelApp.ActiveSheet;
                    //worksheet.Cells[1, "A"] = "Title";
                    //worksheet.Cells[1, "B"] = "URL";

                    //i = 1;
                    //foreach(VideoData video in m_PlaylistTitles)
                    //{
                    //    i++;
                    //    worksheet.Cells[i, "A"] = video.VideoTitle;
                    //    worksheet.Cells[i, "B"] = video.VideoURL;
                    //}
                    //((Range)worksheet.Columns[1]).AutoFit();
                    //((Range)worksheet.Columns[2]).AutoFit();
                    //worksheet.SaveAs($"{Directory.GetCurrentDirectory()}\\{args[1]}.xls");

                    using (StreamWriter sw = new StreamWriter($"{args[1]}.csv"))
                    {
                        sw.WriteLine("Title                              ;URL");
                        foreach (VideoData video in m_PlaylistTitles)
                            sw.WriteLine($"{video.VideoTitle};{video.VideoURL}");
                    }

                }
            }catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }        
    }
}
