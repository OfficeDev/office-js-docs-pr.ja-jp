Office JavaScript API ライブラリには、`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` にある Office JS コンテンツ配信ネットワーク (CDN) を経由してアクセスできます。 アドインの Web ページで Office JavaScript API を使用するには、ページの `<head>` タグにある `<script>` タグに含まれている CDN を参照する必要があります。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。

IntelliSense の入手方法など、Office JavaScript API ライブラリにアクセスする方法の詳細については、「[Office JavaScript API ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)」をご覧ください。