---
title: Office アドインのローカライズ
description: Office JavaScript API を使用して、ロケールを決定し、Office アプリケーションのロケールに基づいて文字列を表示したり、データのロケールに基づいてデータを解釈または表示することができます。
ms.date: 07/21/2020
localization_priority: Normal
ms.openlocfilehash: 2e44c3d28ffc11328e1fae0707c641259f687e8e
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839895"
---
# <a name="localization-for-office-add-ins"></a>Office アドインのローカライズ

Office アドイン に適切なローカライズ方法を任意に選んで実装できます。 JavaScript API と Office アドイン プラットフォームのマニフェスト スキーマには、いくつかの選択肢が用意されています。 Office JavaScript API を使用して、ロケールを決定し、Office アプリケーションのロケールに基づいて文字列を表示したり、データのロケールに基づいてデータを解釈または表示することができます。 マニフェストを使用すると、ロケールに固有なアドイン ファイルの場所と説明的な情報を指定できます。 または、Microsoft Ajax スクリプトを使用して、グローバリゼーションとローカライズをサポートできます。

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>ロケール固有文字列を判別するための JavaScript API の使用

次Office JavaScript API には、アプリケーションおよびデータのロケールと一致する値の表示または解釈をサポートする 2 つのOfficeがあります。

- [Context.displayLanguage][displayLanguage] は、アプリケーションのユーザー インターフェイスのロケール (または言語) Officeします。 次の例では、Office アプリケーションが en-US ロケールまたは fr-FR ロケールを使用しているのか確認し、ロケール固有の案内応答を表示します。

    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }
    ```

- [Context.contentLanguage][contentLanguage] では、データのロケール (または言語) を指定します。[displayLanguage] プロパティを確認する代わりに、最後のコード サンプルを拡張します。`myLanguage` を [contentLanguage] プロパティの値に割り当て、同じコードの残りの部分を使用して、データのロケールに基づいて、あいさつ文を表示します。

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>マニフェストでのローカライズの制御


すべてのOfficeアドインは [、DefaultLocale] 要素とそのマニフェスト内のロケールを指定します。 既定では、Office アドイン プラットフォームと Office クライアント アプリケーションは、Description、DisplayName、IconUrl、HighResolutionIconUrl、[][および SourceLocation]要素の値をすべてのローカルに適用します。 [] [] [] 必要に応じて、これらの 5 つの要素に対して追加ロケールごとに [Override] 子要素を指定することで、特定のロケールの特定の値をサポートできます。 [DefaultLocale]要素と Override 要素の属性の値は `Locale` [、RFC 3066]の「言語の識別のためのタグ」に従って指定されます。 [] 表 1 に、これらの要素のローカライズ サポートを示します。

*表 1.ローカライズのサポート*


|**要素**|**ローカライズのサポート**|
|:-----|:-----|
|[Description]   |指定した各ロケールのユーザーには、AppSource (またはプライベート カタログ) でローカライズされたアドインの説明が表示されます。<br/>Outlook アドインについては、アドインのインストール後に Exchange 管理センター (EAC) に表示される説明が該当します。|
|[DisplayName]   |指定した各ロケールのユーザーには、AppSource (またはプライベート カタログ) でローカライズされたアドインの説明が表示されます。<br/>Outlook アドインについては、アドインのインストール後に [Outlook アドイン] ボタンのラベルおよび EAC に表示される表示名が該当します。<br/>コンテンツ アドインおよび作業ウィンドウ アドインについては、アドインのインストール後にリボンに表示される表示名が該当します。|
|[IconUrl]        |アイコンのイメージは省略可能です。ここで説明したオーバーライドと同じ方法で、特定のカルチャに特定のイメージを指定できます。アイコンを使用し、ローカライズした場合、指定した各ロケールのユーザーには、アドインのローカライズされたアイコン画像が表示されます。<br/>Outlook アドインについては、アドインのインストール後に EAC で表示されるアイコンが該当します。<br/>コンテンツ アドインおよび作業ウィンドウ アドインについては、アドインのインストール後にリボンに表示されるアイコンが該当します。|
|[HighResolutionIconUrl] **重要:** この要素は、アドイン マニフェストのバージョン 1.1 を使用する場合にのみ使用できます。|高解像度のアイコンのイメージは省略可能ですが、指定する場合は、[IconUrl] 要素の後に指定する必要があります。[HighResolutionIconUrl] が指定され、高解像度の dpi をサポートするデバイスにアドインがインストールされている場合、[IconUrl] の値の代わりに [HighResolutionIconUrl] の値が使用されます。<br/>ここで説明したオーバーライドと同じ方法で、特定のカルチャに特定のイメージを指定できます。アイコンを使用し、ローカライズした場合、指定した各ロケールのユーザーには、アドインのローカライズされたアイコン画像が表示されます。<br/>Outlook アドインについては、アドインのインストール後に EAC で表示されるアイコンが該当します。<br/>コンテンツ アドインおよび作業ウィンドウ アドインについては、アドインのインストール後にリボンに表示されるアイコンが該当します。|
|[Resources] **重要:** この要素は、アドイン マニフェストのバージョン 1.1 を使用する場合にのみ使用できます。   |指定した各ロケールのユーザーには、そのロケールのアドイン専用に作成した文字列とアイコン リソースが表示されます。 |
|[SourceLocation]   |指定した各ロケールのユーザーには、そのロケールのアドイン専用にデザインした Web ページが表示されます。 |


> [!NOTE]
> Office がサポートするロケールでのみ、説明と表示名をローカライズできます。 現在の Office のリリースの言語およびロケールの一覧については、「[Office 2013 の言語識別子と OptionState ID 値](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))」を参照してください。


### <a name="examples"></a>例

たとえば、Office アドインで [DefaultLocale] を `en-us` に指定できます。次に示すように、アドインは、[DisplayName] 要素に対して、ロケールが `fr-fr` の [Override] 子要素を指定できます。


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> `de-de` や `de-at` など、言語ファミリ内の複数の領域用にローカライズを行う必要がある場合は、領域ごとに別々の `Override` 要素を使用することをお勧めします。 この場合、言語名を単独で使用する場合、クライアント アプリケーションとプラットフォームのすべての組み合わせOfficeサポート `de` されていません。

このように指定すると、既定ではアドインは `en-us` ロケールを想定します。ほとんどのロケールでは、英語の表示名 "Video player" が表示されます。ただし、クライアント コンピューターのロケールが `fr-fr` の場合は、フランス語の表示名 "Lecteur vidéo" が表示されます。

> [!NOTE]
> 既定のロケールを含め、1 つの言語につき 1 つの override のみを指定できます。 たとえば、既定のロケールが `en-us` の場合、`en-us` の override も指定することはできません。 

次の例では、[Description] 要素に対してロケールのオーバーライドを適用しています。最初に、既定のロケールを `en-us` に指定し、英語の説明を記述しています。その後で、[Override] 要素を使用して、`fr-fr` ロケールに対応するフランス語の説明を記述しています。

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

つまり、アドインでは、既定で `en-us` ロケールを想定します。ほとんどのロケールでは、`DefaultValue` 属性で記述した英語の説明が表示されます。ただし、クライアント コンピューターのロケールが `fr-fr` の場合は、フランス語の説明が表示されます。

次の例では、アドインは、`fr-fr` ロケールとカルチャに対してより適切な別のイメージを指定しています。既定ではイメージ DefaultLogo.png が表示されますが、クライアント コンピューターのロケールが `fr-fr` の場合は、イメージ FrenchLogo.png が表示されます。 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

次の例は、リソースを `Resources` セクションにローカライズする方法を示しています。`ja-jp` カルチャに適したイメージのローカル式を適用しています。

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


[SourceLocation] 要素については、他のロケールをサポートする場合、該当する各ロケール用のソース HTML ファイルを個別に用意する必要があります。指定した各ロケールのユーザーには、そのロケール用にカスタマイズしてデザインした Web ページが表示されます。

Outlook アドインについては、[SourceLocation] 要素もフォーム ファクターに合わせることができます。これにより、対応するフォーム ファクターごとに個別のローカライズされたソース HTML ファイルを指定できます。該当する各設定要素 ([DesktopSettings]、[TabletSettings]、または [PhoneSettings]) に対し、1 つまたは複数の [Override] 子要素を指定できます。次の例は、デスクトップ、タブレット、およびスマートフォンのフォーム ファクターの設定要素を示します。各フォーム ファクターには、既定のロケールを表す HTML ファイルとフランスのロケールを表す HTML ファイルがあります。


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>
</PhoneSettings>
```

## <a name="match-datetime-format-with-client-locale"></a>日付/時刻の形式のクライアント ロケールへの関連付け

[displayLanguage]プロパティを使用して、Officeクライアント アプリケーションのユーザー インターフェイスのロケールを取得できます。 その後、日付と時刻の値を、現在のアプリケーションの現在のロケールと一致する形式でOfficeできます。 その方法の 1 つが、Office アドインがサポートする各ロケールで使用する日付と時刻の表示形式を指定したリソース ファイルを準備するという方法です。 実行時には、アドインはそのリソース ファイルを使用して、[displayLanguage] プロパティから取得したロケールに一致する適切な日付と時刻の形式を判断します。

[contentLanguage]プロパティを使用して、Officeクライアント アプリケーションのデータのロケールを取得できます。 この値に基づいて、日付と時刻の文字列を適切に変換または表示できます。 たとえば、`jp-JP` ロケールでは日付と時刻の値は `yyyy/MM/dd` と表記され、`fr-FR` ロケールでは `dd/MM/yyyy` と表記されます。


## <a name="use-ajax-for-globalization-and-localization"></a>グローバリゼーションとローカライズでの Ajax の使用


Visual Studio で Office アドインを作成する場合, .NET Framework と Ajax を使用してクライアント スクリプト ファイルをグローバライズおよびローカライズできます。

現在のブラウザーのロケール設定に基づいて値を表示するには、Office アドイン向けの JavaScript コード内で [Date](/previous-versions/bb310850(v=vs.140)) および [Number](/previous-versions/bb310835(v=vs.140)) JavaScript 型の拡張と JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) オブジェクトをグローバライズして使用できます。詳細については、「 [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140))」を参照してください。

ローカライズされたリソース文字列をスタンドアロンの JavaScript ファイルに直接埋め込むことで、異なるロケール用のクライアント スクリプト ファイルを作成できます。クライアント スクリプト ファイルは、ブラウザーで設定されるかユーザーが指定できます。サポートされているすべてのロケールに個別のスクリプト ファイルを作成してください。各スクリプト ファイルには、特定のロケール用のリソース文字列を含むオブジェクトを JSON 形式で埋め込みます。ローカライズされた値は、スクリプトがブラウザーで実行されると適用されます。


## <a name="example-build-a-localized-office-add-in"></a>例: ローカライズされた Office アドインの作成

このセクションでは、Office アドイン の説明、表示名、および UI をローカライズする方法の例を示します。 

> [!NOTE]
> 2019 Visual Studioダウンロードするには、IDE Visual Studio [覧ください](https://visualstudio.microsoft.com/vs/)。 インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a>表示または編集用の追加言語を使用できるように Office を構成する

提供されているサンプル コードを実行するには、追加の言語を使用するコンピューターで Office を構成し、メニューとコマンドの表示に使用する言語、編集と校正、あるいはその両方に使用する言語を切り替えてアドインをテストできます。

Office Language Pack を使用して、追加言語をインストールできます。 言語パックの詳細と入手先については、「[Office 2013 の言語オプション](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f)」を参照してください。

Language Accessory Pack をインストールしたら、UI の表示、ドキュメント コンテンツの編集、またはその両方にインストールされた言語を使用するように Office を構成できます。 この記事の例では、言語パック (スペイン語) が適用されている Office のインストールを使用します。

### <a name="create-an-office-add-in-project"></a>Office アドイン プロジェクトの作成

アドイン プロジェクトに 2019 Visual Studio 2019 Officeする必要があります。

> [!NOTE]
> 2019 年 1 月 2019 Visual Studioインストールしていない場合は、ダウンロード手順Visual Studio [IDE](https://visualstudio.microsoft.com/vs/) のページを参照してください。 インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。 Visual Studio 2019 をインストール済みである場合は [、Visual Studio](/visualstudio/install/modify-visual-studio/) インストーラーを使用して、Office/SharePoint 開発ワークロードがインストールされていることを確認します。

1. [**新規プロジェクトの作成**] を選択します。

2. 検索ボックスを使用して、**アドイン** と入力します。 [**Word Web アドイン**] を選択し、[**次へ**] を選択します。

3. プロジェクトに **WorldReadyAddIn という名前を付け、[作成** ] を **選択します**。

4. ソリューションが Visual Studio によって作成され、2 つのプロジェクトが **ソリューション エクスプローラー** に表示されます。 次 **Home.html ファイル** が開Visual Studio。


### <a name="localize-the-text-used-in-your-add-in"></a>アドインに使用されるテキストのローカライズ

別の言語にローカライズする必要があるのは、次の 2 つの領域に表示されるテキストです。

-  **アドインの表示名と説明**。これは、アドインのマニフェスト ファイルのエントリによって制御されています。

-  **アドイン UI**。 アドインの UI に表示される文字列は、JavaScript コードを使用してローカライズできます。たとえば、ローカライズされた文字列を含む別個のリソース ファイルを使用します。

アドインの表示名と説明をローカライズするには

1. **ソリューション エクスプローラー** で、**WorldReadyAddIn**、**WorldReadyAddInManifest** の順に展開し、**WorldReadyAddIn.xml** を選択します。

2. WorldReadyAddInManifest.xml で、[DisplayName] 要素と [Description] 要素を次のコード ブロックに置き換えます。

    > [!NOTE]
    > この例の [DisplayName] 要素および [Description] 要素で使用されているスペイン語にローカライズされた文字列を、別の言語にローカライズされた文字列で置き換えることができます。

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. たとえば、Office 2013 の表示言語を英語からスペイン語に変更してアドインを実行すると、アドインの表示名と説明がローカライズされたテキストで表示されます。

アドインの UI をレイアウトするには

1. Visual Studio の **ソリューション エクスプローラー** で、**Home.html** を選択します。

2. Home.html で `<body>` 要素コンテンツを次の HTML に置き換えて、ファイルを保存します。

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

次の図は、残りの手順を完了してアドインを実行したときにローカライズされたテキストが表示される見出し (h1) 要素と段落 (p) 要素を示しています。

*図 1. アドインの UI*

![セクションが強調表示されたアプリのユーザー インターフェイス](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>ローカライズされた文字列を含むリソース ファイルの追加

JavaScript リソース ファイルには、アドイン UI に使用された文字列が含まれます。 サンプル アドイン UI の HTML には、あいさつ文を表示する `<h1>` 要素、およびユーザーにアドインを紹介する `<p>` 要素が含まれます。 

見出しと段落のローカライズされた文字列を有効にするには、文字列を別個のリソース ファイルに置きます。このリソース ファイルにより、ローカライズされた文字列の各セット用に個別の JavaScript Object Notation (JSON) オブジェクトを格納する JavaScript オブジェクトが作成されます。また、指定したロケールに対する適切な JSON オブジェクトを取得するためのメソッドも提供されます。

アドイン プロジェクトにリソース ファイルを追加するには

1. Visual Studio の **ソリューション エクスプローラー** で、**WorldReadyAddInWeb** プロジェクトを右クリックして **[追加]** > **[新しい項目]** を選択します。 

2. **[新しい項目の追加]** ダイアログ ボックスで **[JavaScript ファイル]** を選択します。

3. ファイル名として「**UIStrings.js**」と入力して、**[追加]** を選択します。

4. 次のコードを UIStrings.js ファイルに追加して、ファイルを保存します。

    ```js
    /* Store the locale-specific strings */

    var UIStrings = (function ()
    {
        "use strict";

        var UIStrings = {};

        // JSON object for English strings
        UIStrings.EN =
        {
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."
        };

        // JSON object for Spanish strings
        UIStrings.ES =
        {
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };

        UIStrings.getLocaleStrings = function (locale)
        {
            var text;

            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }

            return text;
        };

        return UIStrings;
    })();
    ```

UIStrings.js リソース ファイルで、アドインの UI のローカライズされた文字列を含むオブジェクト **UIStrings** を作成します。

### <a name="localize-the-text-used-for-the-add-in-ui"></a>アドインの UI に使用するテキストのローカライズ

アドインでリソース ファイルを使用するには、リソース ファイルのスクリプト タグを Home.html に追加する必要があります。Home.html が読み込まれると、UIStrings.js が実行され、文字列の取得に使用する **UIStrings** オブジェクトをコードで利用できるようになります。コードで **UIStrings** を利用できるようにするには、Home.html の head タグに次の HTML を追加します。

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

これで、**UIStrings** オブジェクトを使用してアドインの UI の文字列を設定できるようになりました。

Office クライアント アプリケーションのメニューおよびコマンドでの表示に使用する言語に基づいてアドインのローカライズを変更する場合は **、Office.context.displayLanguage** プロパティを使用してその言語のロケールを取得します。 たとえば、アプリケーション言語がメニューやコマンドの表示にスペイン語を使用している場合 **、Office.context.displayLanguage** プロパティは言語コード es-ES を返します。

ドキュメント コンテンツの編集に使用されている言語に基づいてアドインのローカライズを変更する場合は **、Office.context.contentLanguage** プロパティを使用してその言語のロケールを取得します。 たとえば、アプリケーション言語がドキュメント コンテンツの編集にスペイン語を使用している場合 **、Office.context.contentLanguage** プロパティは言語コード es-ES を返します。

アプリケーションが使用している言語を確認した後 **、UIStrings** を使用して、アプリケーションの言語に一致するローカライズされた文字列のセットを取得できます。

Home.js ファイルのコードを次のコードで置き換えます。 このコードは、アプリケーションの表示言語またはアプリケーションの編集言語に基づいて、Home.html の UI 要素で使用される文字列を変更する方法を示しています。

> [!NOTE]
> 編集で使用した言語にアドインのローカライズを変更して切り替えるには、コード行 `var myLanguage = Office.context.contentLanguage;` をコメント解除し、コード行 `var myLanguage = Office.context.displayLanguage;` をコメント化します。

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the Office application.
            var myLanguage = Office.context.displayLanguage;
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a>ローカライズされたアドインのテスト

ローカライズされたアドインをテストするには、Office アプリケーションで表示または編集に使用する言語を変更し、アドインを実行します。

アドインで表示または編集に使用される言語を変更するには

1. Word で **[ファイル]**、**[オプション]**、**[言語]** の順に選択します。 次の図に、[言語] タブが開かれている **[Word のオプション]** ダイアログ ボックスを示します。

    *図 2. Word の [オプション] ダイアログ ボックスの言語オプション*

    ![Word の [オプション] ダイアログ](../images/office15-app-how-to-localize-fig04.png)

2. **[表示言語の選択]** で表示する言語 (スペイン語など) を選択して、上向き矢印を選択してスペイン語をリストの最初の位置に移動します。 または、編集に使用する言語を変更するには、[編集言語の選択] で編集に使用する言語 (スペイン語など) を選択し、[既定に設定] を選択 **します**。

3. **[OK]** をクリックして選択内容を確認し、Word を閉じます。

4. コマンド **バーで F5** キー Visual Studioサンプル アドインを実行するか、メニュー バーから [デバッグ開始  >  **デバッグ**] を選択します。

5. Word で **[ホーム]**、**[作業ウィンドウを表示]** の順に選択します。

実行すると、次の図に示すように、アドイン UI の文字列がアプリケーションで使用される言語に合わせて変更されます。


*図 3. ローカライズされたテキストが表示されたアドインの UI*

![UI テキストがローカライズされたアプリ](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>関連項目

- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [Office 2013 の言語識別子と OptionState ID 値](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[説明]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displaylanguage
[contentLanguage]:       /javascript/api/office/office.context#contentlanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066