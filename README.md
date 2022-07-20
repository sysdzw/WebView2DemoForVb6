# WebView2DemoForVb6
WebView2Demo for vb6

## 步骤：
1）先执行RegisterRC6inPlace.vbs，注册rc6.dll
2）双击运行MicrosoftEdgeWebview2Setup.exe安装webview2环境。


## 实现原理：
本浏览器依托于大名鼎鼎的VBrichclient，大神把微软发布的webview2封装到了rc6.dll中，使得其他软件可以调用。

用起来还是有点问题，比如拦截新窗口打开的网页直接在WV_NewWindowRequested中用`WV.Navigate URI`打开会很慢很慢，大概要10秒左右，然后我借用定时器timer1，在WV_NewWindowRequested过程外面执行打开，这样才快一点，但是浏览器的路由发生的问题，比如无法正常后退到上一个网页了。如果谁有解决方案的麻烦共享下。

QQ171977759
