Office JavaScript API には、2 つの異なるモデルがあります。

- **アプリケーション固有** API では、特定の Office アプリケーションにネイティブなオブジェクトを操作するために使用できる、厳密に型指定されたオブジェクトが提供されます。 たとえば、Excel JavaScript API を使用して、ワークシート、範囲、テーブル、グラフなどにアクセスすることができます。 アプリケーション固有 API は現在、次の Office アプリケーション用に使用できます。

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    この API モデルでは [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) が使用され、Office アプリケーションに送信する各要求で複数の操作を指定することが可能です。 この方法によるバッチ操作を行うと、Web 上の Office アプリケーションのパフォーマンスが大幅に向上します。 アプリケーション固有の API は Office 2016 で導入されました。Office 2013 の操作には使用できません。

    > [!NOTE]
    > [Visio](../reference/overview/visio-javascript-reference-overview.md) にはホスト固有の API もありますが、それを SharePoint Online ページでのみ使用して、ページに埋め込まれている Visio 図面を操作できます。 Visio では Office Web アドインはサポートされていません。

    この API モデルの詳細については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。

- **共通 API** を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。 この API モデルでは [コールバック](https://developer.mozilla.org/docs/Glossary/Callback_function) が使用され、Office アプリケーションに送信する各要求で 1 つの操作のみを指定できます。 共通 API は Office 2013 で導入されました。Office 2013 以降の操作に使用できます。 Outlook、PowerPoint、Project を操作するための API を含む、共通 API オブジェクト モデルの詳細については、「[共通 JavaScript API オブジェクト モデル](../develop/office-javascript-api-object-model.md)」を参照してください。

> [!NOTE]
>[共有ランタイム](../testing/runtimes.md#shared-runtime)のないカスタム関数は、計算の実行に優先順位を付ける [JavaScript 専用ランタイム](../testing/runtimes.md#javascript-only-runtime)で実行されます。 これらの関数は、少し異なるプログラミング モデルを使用します。
