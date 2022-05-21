---
title: Office のホストと API の要件を指定する
description: アドインが期待どおりに動作するようにOfficeアプリケーションと API 要件を指定する方法について説明します。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60ad00c918b04b6f12ecb6eec6c40772448b2ab8
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628048"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API 要件を指定する

Office アドインは、特定のOffice アプリケーション (Office ホストとも呼ばれます) または Office JavaScript API の特定のメンバー (office.js) に依存する場合があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。
- 一部のバージョンのOfficeでのみ使用できるOffice JavaScript API を使用します。 たとえば、Excel 2016 の 1 回限りの購入バージョンでは、Office JavaScript ライブラリ内のすべてのExcel関連 API がサポートされているわけではありません。

このような状況では、アドインが実行できないOffice アプリケーションまたはOffice バージョンにアドインがインストールされないようにする必要があります。

また、Office アプリケーションとOffice バージョンに基づいて、アドインのどの機能をユーザーに表示するかを制御するシナリオもあります。 次の 2 つの例を示します。

- アドインには、テキスト操作など、Word とPowerPointの両方で役立つ機能がありますが、スライド管理機能など、PowerPointでのみ意味を持つ追加機能がいくつかあります。 アドインが Word で実行されている場合は、PowerPointのみの機能を非表示にする必要があります。
- アドインには、サブスクリプション Excelなど、Office アプリケーションの一部のバージョンでサポートされているが、1 回限りの購入Excel 2016など、他のバージョンではサポートされていない、Office JavaScript API メソッドが必要な機能があります。 ただし、アドインには、Excel 2016でサポート *されている* javaScript API メソッドOffice必要とするその他の機能があります。 このシナリオでは、アドインをExcel 2016にインストールできるようにする必要がありますが、サポートされていない方法を必要とする機能は、Excel 2016のユーザーから非表示にする必要があります。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> Office アドインが現在サポートされている場所の概要については、「[Office クライアント アプリケーションとOffice アドインのプラットフォームの可用性](/javascript/api/requirement-sets)」ページを参照してください。

> [!TIP]
> この記事で説明するタスクの多くは、ツールを使用してアドイン プロジェクトを作成するときに、アドイン プロジェクト全体または部分的に実行されます。たとえば、[Office アドイン用の Yeoman ジェネレーターや、Visual Studio](yeoman-generator-overview.md)のOffice アドイン テンプレートの 1 つなどです。 このような場合は、タスクが完了したことを確認する必要があることを意味として解釈してください。

## <a name="use-the-latest-office-javascript-api-library"></a>最新のOffice JavaScript API ライブラリを使用する

アドインは、コンテンツ配信ネットワーク (CDN) から最新バージョンの Office JavaScript API ライブラリを読み込む必要があります。 これを行うには、アドインが開く最初の HTML ファイルに次 `script` のタグがあることを確認します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>アドインをホストできるOfficeアプリケーションを指定する

既定では、アドインは、指定されたアドインの種類 (つまり、メール、作業ウィンドウ、またはコンテンツ) でサポートされているすべてのOffice アプリケーションにインストールできます。 たとえば、作業ウィンドウ アドインは、Access、Excel、OneNote、PowerPoint、Project、Word で既定でインストールできます。 

アドインがOffice アプリケーションのサブセットに確実にインストールできるようにするには、マニフェストの [Hosts](/javascript/api/manifest/hosts) 要素と [Host](/javascript/api/manifest/host) 要素を使用します。

たとえば、次の **ホスト** と **ホスト** の宣言では、アドインをExcel on the web、Windows、iPadを含むExcelの任意のリリースにインストールできますが、他のOffice アプリケーションにはインストールできないことを指定します。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

**Hosts** 要素には 1 つ以上の **Host** 要素を含めることができます。 アドインをインストール可能にするOffice アプリケーションごとに個別の **Host** 要素が存在する必要があります。 属性は `Name` 必須であり、次のいずれかの値に設定できます。

| 名前          | Office クライアント アプリケーション                     | 使用可能なアドインの種類 |
|:--------------|:-----------------------------------------------|:-----------------------|
| データベース      | Access Web アプリ                                | 作業ウィンドウ              |
| 文書      | Word on the web、Windows、Mac、iPad            | 作業ウィンドウ              |
| Mailbox       | Outlook on the web、Windows、Mac、Android、iOS | メール                   |
| Notebook      | OneNote on the web                             | 作業ウィンドウ、コンテンツ     |
| Presentation  | PowerPoint on the web、Windows、Mac、iPad      | 作業ウィンドウ、コンテンツ     |
| Project       | Windows での Project                             | 作業ウィンドウ              |
| ブック      | Excel on the web、Windows、Mac、iPad           | 作業ウィンドウ、コンテンツ     |

> [!NOTE]
> Office アプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行されます。 通常、アドインの実行に使用できるプラットフォームを指定することはできません。 たとえば、指定`Workbook`した場合は、Excel on the webとWindowsの両方を使用してアドインを実行できます。 ただし、指定`Mailbox`した場合、[モバイル拡張ポイント](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface)を定義しない限り、アドインOutlookモバイル クライアントでは実行されません。

> [!NOTE]
> アドイン マニフェストを複数の種類 (メール、作業ウィンドウ、またはコンテンツ) に適用することはできません。 つまり、アドインをOutlookと他のOffice アプリケーションのいずれかにインストールできるようにする場合は、*2 つの* アドインを作成する必要があります。1 つはメール タイプ マニフェスト、もう 1 つは作業ウィンドウまたはコンテンツ タイプ マニフェストを持つアドインです。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>アドインをホストできるOfficeバージョンとプラットフォームを指定する

Officeバージョンとビルド、またはアドインをインストール可能にするプラットフォームを明示的に指定することはできません。また、アドインで使用するアドイン機能のサポートが新しいバージョンまたはプラットフォームに拡張されるたびにマニフェストを修正する必要があるため、必要はありません。 代わりに、アドインに必要な API をマニフェストで指定します。 Officeは、API をサポートしていないOfficeバージョンとプラットフォームの組み合わせにアドインがインストールされないようにし、アドインが **マイ アドイン** に表示されないようにします。

> [!IMPORTANT]
> 基本マニフェストを使用して、アドインが重要な値である必要がある API メンバーをまったく指定します。 アドインが一部の機能に API を使用しているが、API を必要としないその他の便利な機能がある場合は、API をサポートしていないがそれらの組み合わせに対して少ないエクスペリエンスを提供するプラットフォームとOfficeバージョンの組み合わせにインストールできるように、アドインを設計する必要があります。 詳細については、「 [代替エクスペリエンスの設計」を](#design-for-alternate-experiences)参照してください。

### <a name="requirement-sets"></a>要件セット

アドインで必要な API を指定するプロセスを簡略化するために、ほとんどの API を *要件セット* にグループ化Office。 [共通 API オブジェクト モデル](understanding-the-javascript-api-for-office.md#api-models)の API は、それらがサポートする開発機能によってグループ化されます。 たとえば、テーブル バインドに接続されているすべての API は、"TableBindings 1.1" という要件セットにあります。 [アプリケーション固有のオブジェクト モデル](understanding-the-javascript-api-for-office.md#api-models)の API は、運用環境のアドインで使用するためにリリースされた時点でグループ化されます。

要件セットはバージョン管理されています。 たとえば、 [ダイアログ ボックス](../develop/dialog-api-in-office-add-ins.md) をサポートする API は、要件セット DialogApi 1.1 にあります。 作業ウィンドウからダイアログへのメッセージングを有効にする追加の API がリリースされると、DialogApi 1.2 と DialogApi 1.1 のすべての API にグループ化されました。 *要件セットの各バージョンは、以前のすべてのバージョンのスーパーセットです。*

要件セットのサポートは、Office アプリケーション、Office アプリケーションのバージョン、およびそれが実行されているプラットフォームによって異なります。 たとえば、DialogApi 1.2 は、Office 2021以前の 1 回限りの購入バージョンのOfficeではサポートされていませんが、DialogApi 1.1 は 2013 Officeに戻るすべての 1 回限りの購入バージョンでサポートされています。 アドインを、使用する API をサポートするプラットフォームとOffice バージョンのすべての組み合わせにインストールできるようにする必要があるため、アドインで必要な各要件セットの *最小* バージョンをマニフェストで常に指定する必要があります。 これを行う方法の詳細については、この記事の後半で説明します。

> [!TIP]
> 要件セットのバージョン管理の詳細については、「[Office要件セットの可用性](office-versions-and-requirement-sets.md#office-requirement-sets-availability)」を参照し、要件セットの完全な一覧と各要件セットの API に関する情報については、[アドイン要件セットOffice](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)から開始します。 ほとんどのOffice.js API のリファレンス トピックでは、それらが属する要件セット (存在する場合) も指定されています。

> [!NOTE]
> 一部の要件セットには、マニフェスト要素も関連付けられます。 この事実がアドイン デザインに関連する場合の詳細については、「 [VersionOverrides 要素での要件の指定](#specify-requirements-in-a-versionoverrides-element) 」を参照してください。

#### <a name="apis-not-in-a-requirement-set"></a>要件セットに含まれていない API

アプリケーション固有のモデル内のすべての API は要件セットに含まれていますが、Common API モデルの一部は要件セットに含まれていません。 また、アドインで必要な場合に、マニフェストでこれらのセットレス API の 1 つを指定する方法もあります。 詳細については、この記事で後述します。

### <a name="requirements-element"></a>Requirements 要素

[Requirements](/javascript/api/manifest/requirements) 要素とその子要素[セット](/javascript/api/manifest/sets)と[メソッド](/javascript/api/manifest/methods)を使用して、アドインをインストールするためにOffice アプリケーションでサポートする必要がある最小要件セットまたは API メンバーを指定します。 

Office アプリケーションまたはプラットフォームが **Requirements** 要素で指定された要件セットまたは API メンバーをサポートしていない場合、アドインはそのアプリケーションまたはプラットフォームでは実行されないため、**マイ アドイン** には表示されません。

> [!NOTE]
> **Requirements** 要素は、Outlook アドインを除くすべてのアドインで省略可能です。ルート要素の`xsi:type`属性が存在する場合は`MailApp`、アドインが **必要とするメールボックス** 要件セットの最小バージョンを指定する Requirements 要素が必要です。`OfficeApp` 詳細については、[JavaScript API 要件セットOutlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)を参照してください。

次のコード例は、次をサポートするすべてのOffice アプリケーションでインストール可能なアドインを構成する方法を示しています。

-  `TableBindings` 要件セット。最小バージョンは "1.1" です。
-  `OOXML` 要件セット。最小バージョンは "1.1" です。
-  `Document.getSelectedDataAsync` メソッド。

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```
この例については、次の点に注意してください。

- **Requirements** 要素には、**Sets** 子要素と **Methods** 子要素が含まれています。
- **Sets** 要素には、1 つ以上の **Set 要素を** 含めることができます。 `DefaultMinVersion`は、すべての子 Set 要素の既定値`MinVersion`**を** 指定します。
- [Set](/javascript/api/manifest/set) 要素は、アドインをインストール可能にするためにOffice アプリケーションがサポートする必要がある要件セットを指定します。 属性は `Name` 、要件セットの名前を指定します。 要件 `MinVersion` セットの最小バージョンを指定します。 `MinVersion` は、親 Sets 内の `DefaultMinVersion` 属性の値をオーバーライド **します**。
- **Methods** 要素には、1 つ以上の [メソッド要素を](/javascript/api/manifest/method)含めることができます。 Outlook アドインで **Methods** 要素を使用することはできません。
- **Method** 要素は、アドインをインストール可能にするために、Office アプリケーションがサポートする必要がある個々のメソッドを指定します。 属性は `Name` 必須であり、その親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="design-for-alternate-experiences"></a>代替エクスペリエンスの設計

Office アドイン プラットフォームが提供する機能拡張機能は、次の 3 種類に分けることができます。

- アドインがインストールされた直後に使用できる機能拡張。 この種の機能を利用するには、マニフェストで [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素を構成します。 この種の機能の例として、カスタム リボン ボタンとメニューである [アドイン コマンド](../design/add-in-commands.md)があります。
- アドインが実行されていて、Office.js JavaScript API で実装されている場合にのみ使用できる機能拡張機能。たとえば、 [ダイアログ ボックスなどです](../develop/dialog-api-in-office-add-ins.md)。
- 拡張機能機能は、実行時にのみ使用できますが、 **VersionOverrides** 要素でOffice.js JavaScript と構成の組み合わせで実装されます。 これらの例は、[カスタム関数](../excel/custom-functions-overview.md)、[シングル サインオン](sso-in-office-add-ins.md)、およびカスタム [コンテキスト タブExcelです](../design/contextual-tabs.md)。

アドインが機能の一部に特定の機能拡張機能を使用しているが、機能拡張機能を必要としないその他の便利な機能がある場合は、拡張機能機能をサポートしないプラットフォームとOfficeバージョンの組み合わせにインストールできるようにアドインを設計する必要があります。 これらの組み合わせに対する貴重なエクスペリエンスを提供できます。 

この設計は、機能拡張機能の実装方法に応じて異なります。 

- JavaScript で完全に実装される機能については、 [メソッドと要件セットのサポートに関するランタイム チェックに関する](#runtime-checks-for-method-and-requirement-set-support)ページを参照してください。
- **VersionOverrides** 要素を構成する必要がある機能については、「[VersionOverrides 要素での要件の指定」を](#specify-requirements-in-a-versionoverrides-element)参照してください。

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>メソッドと要件セットのサポートのランタイム チェック 

実行時にテストして、ユーザーのOffice[が isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) メソッドで要件セットをサポートしているかどうかを検出します。 要件セットの名前と最小バージョンをパラメーターとして渡します。 要件セットがサポートされている場合は、 `isSetSupported` **true** を返します。 次のコードは一例です。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
このコードについては、以下の点に注意してください。

- 最初のパラメーターが必要です。 これは、要件セットの名前を表す文字列です。 利用できる要件セットの詳細については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)」を参照してください。
- 2 番目のパラメーターは省略可能です。 これは、ステートメント内`if`のコードを実行するためにOffice アプリケーションがサポートする必要がある最小要件セット バージョンを指定する文字列です (例: "**1.9**")。 使用しない場合は、バージョン "1.1" が想定されます。

> [!WARNING]
> メソッドを `isSetSupported` 呼び出す場合、2 番目のパラメーターの値 (指定されている場合) は数値ではなく文字列にする必要があります。 JavaScript パーサーは 1.1 や 1.10 などの数値を区別できませんが、"1.1" や "1.10" などの文字列値では区別できます。

次の表は、アプリケーション固有の API モデルの要件セット名を示しています。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

Common API モデル要件セットの 1 つでメソッドを使用する例を次に示します。

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> これらのアプリケーションのメソッドと要件セットは`isSetSupported`、CDNの最新のOffice.js ファイルで使用できます。 CDNからOffice.jsを使用しない場合、古いバージョンのライブラリを使用している場合は、アドインで`isSetSupported`例外が生成される可能性があります。 詳細については、「[最新のOffice JavaScript API ライブラリを使用](#use-the-latest-office-javascript-api-library)する」を参照してください。

アドインが要件セットの一部ではないメソッドに依存している場合は、次のコード例に示すように、ランタイム チェックを使用して、そのメソッドがOffice アプリケーションによってサポートされているかどうかを判断します。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例では、Office アプリケーションでサポート`document.setSelectedDataAsync`されているかどうかを確認します。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>VersionOverrides 要素で要件を指定する

[VersionOverrides](/javascript/api/manifest/versionoverrides) 要素は、アドイン コマンド (カスタム リボン ボタンやメニュー) などのアドインのインストール直後に使用できる必要がある機能をサポートするために、主にマニフェスト スキーマに追加されましたが、排他的ではありません。 アドイン マニフェストOffice解析するときに、これらの機能について把握しておく必要があります。 

アドインでこれらの機能のいずれかを使用しているが、アドインは価値があり、機能をサポートしていないOfficeバージョンでもインストール可能である必要があるとします。 このシナリオでは、基本`OfficeApp`要素の子としてではなく **、VersionOverrides** 要素自体の子として含める [Requirements](/javascript/api/manifest/requirements) 要素 (およびその子 [Sets](/javascript/api/manifest/sets) および [Methods](/javascript/api/manifest/methods) 要素) を使用して機能を特定します。 これを行うと、Officeアドインのインストールが許可されますが、機能がサポートされていないOfficeバージョンの **VersionOverrides** 要素の子要素は無視Office。

具体的には、基本マニフェストの要素 (**Hosts** 要素など) をオーバーライドする **VersionOverrides** の子要素は無視され、代わりに基本マニフェストの対応する要素が使用されます。 ただし、 **VersionOverrides** には、基本マニフェストの設定をオーバーライドするのではなく、実際に追加の機能を実装する子要素が存在する可能性があります。 2 つの例を次に示します`WebApplicationInfo``EquivalentAddins`。 プラットフォームとバージョンのOfficeが対応する機能をサポートしている場合、**VersionOverrides** のこれらの部分は無視 *されません*。  

**Requirements** 要素の子孫要素の詳細については、この記事の前半の [「Requirements 要素](#requirements-element)」を参照してください。

次に例を示します。

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> 要件をサポートしていないプラットフォームとバージョンの組み合わせ *では、要件* を *必要としない機能を呼び出す* アドイン コマンドもインストールされないため、**VersionOverrides** で **Requirements 要素を** 使用する前に注意してください。 たとえば、2 つのカスタム リボン ボタンを持つアドインを考えてみましょう。 そのうちの 1 つは、要件セット **ExcelApi 1.4** (以降) で使用できる JavaScript API Office呼び出します。 他の呼び出し API は **、ExcelApi 1.9** (以降) でのみ使用できます。 **VersionOverrides** に **ExcelApi 1.9** の要件を設定した場合、1.9 がサポートされていない場合 *、どちらのボタンも* リボンに表示されません。 このシナリオでは、 [メソッドと要件セットのサポートに関するランタイム チェック](#runtime-checks-for-method-and-requirement-set-support)で説明されている手法を使用することをお勧めします。 2 番目のボタンによって呼び出されるコードは、最初に **ExcelApi 1.9** のサポートをチェックするために使用`isSetSupported`します。 サポートされていない場合、このコードでは、アドインのこの機能は自分のバージョンのOfficeでは使用できないことを示すメッセージがユーザーに表示されます。 

> [!TIP]
> 基本マニフェストに既に表示されている **VersionOverrides** で **要件** 要素を繰り返しても意味がありません。 基本マニフェストで要件が指定されている場合、アドインは要件がサポートされていない場所にインストールできないため、**Office VersionOverrides** 要素を解析しません。 

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](add-in-manifests.md)
- [Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
