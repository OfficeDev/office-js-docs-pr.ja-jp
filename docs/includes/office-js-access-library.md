<span data-ttu-id="53095-101">Office JavaScript API ライブラリには、`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` にある Office JS コンテンツ配信ネットワーク (CDN) を経由してアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="53095-101">The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="53095-102">アドインの Web ページで Office JavaScript API を使用するには、ページの `<head>` タグにある `<script>` タグに含まれている CDN を参照する必要があります。</span><span class="sxs-lookup"><span data-stu-id="53095-102">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="53095-103">プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。</span><span class="sxs-lookup"><span data-stu-id="53095-103">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

<span data-ttu-id="53095-104">IntelliSense の入手方法など、Office JavaScript API ライブラリにアクセスする方法の詳細については、「[Office JavaScript API ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="53095-104">For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>