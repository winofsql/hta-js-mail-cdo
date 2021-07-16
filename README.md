# HTML アプリケーション : CDO によるメール送信
```javascript
var cdo = new ActiveXObject("CDO.Message");
var server = "smtp.lolipop.jp";
var port = 465;
var user = "ユーザ";
var from = "メールアドレス";
var pass = "パスワード";

```
```html
<meta http-equiv="x-ua-compatible" content="ie=edge">
<meta charset="utf-8">
```

## Visual Studio Code の settings.json
```javascript
"files.associations": {
   "*.hta": "html"
}
```
## mshta.exe は 32ビットアプリケーションです
"C:\Windows\SysWOW64\mshta.exe"
