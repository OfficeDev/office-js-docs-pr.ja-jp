Office JavaScript API には、2 つの異なるモデルがあります。

- **ホスト固有** API では、特定の Office アプリケーションにネイティブなオブジェクトを操作するために使用できる、厳密に型指定されたオブジェクトが提供されます。 たとえば、Excel JavaScript API を使用して、ワークシート、範囲、テーブル、グラフなどにアクセスすることができます。 ホスト固有 API は現在、次のホスト用に使用できます。

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)

    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    この API モデルでは [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) が使用され、Office ホストに送信する各要求で複数の操作を指定することが可能です。 この方法によるバッチ操作を行うと、Office on the web アプリケーションのパフォーマンスが大幅に向上します。 ホスト固有の API は Office 2016 で導入されました。Office 2013 の操作には使用できません。

- **共通 API** を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。 この API モデルでは [Callback](https://developer.mozilla.org/docs/Glossary/Callback_function) が使用され、Office ホストに送信する各要求で指定できる操作は、1 つのみです。 共通 API は Office 2013 で導入されました。Office 2013 以降の操作に使用できます。 Outlook と PowerPoint を操作するための API を含む、共通 API オブジェクト モデルの詳細については、「[共通 JavaScript API オブジェクト モデル](../develop/office-javascript-api-object-model.md)」を参照してください。

> [!NOTE]
> Excel のカスタム関数の場合は、計算の実行を優先する独自のランタイム内で実行されるため、少し異なるプログラミング モデルが使用されます。 詳細については、「[カスタム関数のアーキテクチャ](../excel/custom-functions-architecture.md)」を参照してください。