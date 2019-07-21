---
ms.date: 07/10/2019
description: JavaScript を使用して Excel でカスタム関数を作成する。
title: Excel でカスタム関数を作成する
localization_priority: Priority
ms.openlocfilehash: c5b31b494d7b22112e36e245603f58748559bed5
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771405"
---
# <a name="create-custom-functions-in-excel"></a>Excel でカスタム関数を作成する 

開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。 この記事では、Excel でカスタム関数を作成する方法について説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次のアニメーション画像は、JavaScript または Typescript で作成した関数を呼び出すブックを示しています。 この例では、カスタム関数 `=MYFUNCTION.SPHEREVOLUME` は球の体積を計算します。

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

`=MYFUNCTION.SPHEREVOLUME` カスタム関数は次のコードにより定義されます。

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> この記事で後述する「[既知の問題](#known-issues)」セクションで、カスタム関数の現状の制限事項を記載します。

## <a name="how-a-custom-function-is-defined-in-code"></a>コードでカスタム関数を定義する方法

[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数、作業ウィンドウ、およびアドイン全体をこのジェネレーターが作成します。 このため、カスタム関数に重要なファイルに注意を集中できます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>または<br/>**./src/functions/functions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含みます。 |
| **./src/functions/functions.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | アドイン内のすべてのカスタム関数の名前空間と、この表で前述した JavaScript ファイルと HTML ファイルの位置を指定します。 また、作業ウィンドウ ファイルやコマンド ファイルなど、アドインで使用する可能性のある他のファイルの位置もリストされます。 |

### <a name="script-file"></a>スクリプト ファイル

スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義するコードと関数を定義するコメントが含まれています。

`add` カスタム関数は次のコードにより定義されます。 コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。 必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。 さらに、お気付きのように `first` と `second` の 2 つのパラメーターが宣言されており、その後にそれらの `description` プロパティが記述されます。 最後に `returns` の説明が記述されます。 カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)」を参照してください。

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

カスタム関数のランタイムの読み込みを制御する **functions.html** ファイルは、カスタム関数の現在の CDN にリンクしていなければならないことに注意してください。 最新バージョンの Yo Office ジェネレーターを使用して作成されたプロジェクトは、正しい CDN を参照します。 2019 年 3 月以前の古いカスタム関数のプロジェクトを改良する場合は、以下のコードを **functions.html** ページにコピーする必要があります。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a>マニフェスト ファイル

カスタム関数 (Yo Office ジェネレーターが作成するプロジェクトでは **./manifest.xml**) を定義するアドインの XML マニフェスト ファイルは、アドイン内のすべてのカスタム関数の名前空間と、 JavaScript、JSON、および HTML の場所を指定します。

次の基本的な XML マークアップは、カスタム関数を有効にするアドインのマニフェストに含める必要がある要素`<ExtensionPoint>` と `<Resources>` の例を示しています。 Yo Office ジェネレーターを使用する場合、生成されたカスタム関数ファイルには、さらに複雑なマニフェスト ファイルが格納されます。こちらの[Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)で比較できます。

> [!NOTE] 
> カスタム関数のJavaScript、JSON、HTML ファイルのマニフェスト ファイルで指定した URL はだれでもアクセスでき、同じサブドメインを持つ必要があります。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Excel の関数は、XML マニフェスト ファイルで指定された名前空間が接頭辞として付加されます。 関数の名前空間は、関数名の前に付けられ、ピリオドで区切られます。 例えば、Excel ワークシートのセル内で、`ADD42` 関数を呼び出すためには、`=CONTOSO.ADD42` と入力します。これは、`CONTOSO` が名前空間で、`ADD42` が JSON ファイルで指定された関数の名前だからです。 名前空間は、会社またはアドインの識別子としての使用を目的としています。 名前空間にはアルファベットとピリオドのみを含めることが出来ます。

## <a name="coauthoring"></a>共同編集

Excel on the web と Office 365 サブスクリプションに接続している Windows の場合は、ドキュメントの共同編集を行うことができます。また、この機能でカスタム関数を使用できます。 ブックでカスタム関数を使用している場合、仕事仲間はカスタム関数のアドインを読み込むように要求されます。 双方がアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。

共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。

## <a name="known-issues"></a>既知の問題

既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。

## <a name="next-steps"></a>次の手順

カスタム関数を試してみましょう。 もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。

独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。 独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。

カスタム関数の機能の詳細について読む準備はできましたか? [カスタム関数のアーキテクチャ](custom-functions-architecture.md)の概要をご覧ください。

## <a name="see-also"></a>関連項目 
* [カスタム関数の要件](custom-functions-requirement-sets.md)
* [名前付けのガイドライン](custom-functions-naming.md)
* [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](make-custom-functions-compatible-with-xll-udf.md)
