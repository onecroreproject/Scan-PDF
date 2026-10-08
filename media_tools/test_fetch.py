import urllib.request

req = urllib.request.Request("https://twitter.com/x/status/2107760960136004062", headers={'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'})
try:
    with urllib.request.urlopen(req, timeout=10) as resp:
        html = resp.read().decode('utf-8')
        with open("media_tools/twitter_test.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("Success, wrote to twitter_test.html")
except Exception as e:
    print("Error:", e)
