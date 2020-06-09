---
ms.date: 05/17/2020
description: Office アドインの Excel カスタム関数を作成する
title: Excel でカスタム関数を作成する
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 4f8416b9058def9dcb4998fb2f31684b59276ac4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609283"
---
# <a name="create-custom-functions-in-excel"></a>Excel でカスタム関数を作成する

開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。

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

[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Excel カスタム関数アドインプロジェクトを作成する場合は、関数と作業ウィンドウを制御するファイルを作成します。 このため、カスタム関数に重要なファイルに注意を集中できます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>または<br/>**./src/functions/functions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含みます。 |
| **./src/functions/functions.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | カスタム関数で使用する複数のファイルの場所を指定します。これには、カスタム関数 JavaScript、JSON、HTML ファイルなどがあります。 また、作業ウィンドウファイルやコマンドファイルの場所の一覧を示し、カスタム関数が使用する必要があるランタイムを指定します。 |

### <a name="script-file"></a>スクリプト ファイル

スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義するコードと関数を定義するコメントが含まれています。

`add` カスタム関数は次のコードにより定義されます。 コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。 必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。 次に、2つのパラメーターが宣言され、その `first` `second` 後にプロパティが続き `description` ます。 最後に `returns` の説明が記述されます。 カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)」を参照してください。

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

### <a name="manifest-file"></a>マニフェスト ファイル

ユーザー設定の XML マニフェストファイルは、Yo Office ジェネレーターによって作成されたプロジェクト内のカスタム関数 (**./manifest¥ .xml** ) を定義します。

- カスタム関数の名前空間を定義します。 ユーザーが自分の関数をアドインの一部として識別できるようにするために、名前空間がカスタム関数に追加されています。
- `<ExtensionPoint>` `<Resources>` カスタム関数マニフェストに固有のおよび要素を使用します。 これらの要素には、JavaScript、JSON、および HTML ファイルの場所に関する情報が含まれています。
- カスタム関数に対して使用するランタイムを指定します。 共有ランタイムでは、関数と作業ウィンドウとの間でデータを共有できるため、別のランタイムに特に必要性がある場合を除き、常に共有ランタイムを使用することをお勧めします。

Yo Office ジェネレーターを使用してファイルを作成する場合は、共有ランタイムを使用するようにマニフェストを調整することをお勧めします。これは、これらのファイルの既定値ではないためです。 マニフェストを変更するには、「 [Excel アドインを構成する](./configure-your-add-in-to-use-a-shared-runtime.md)」の手順に従って、共有されている JavaScript ランタイムを使用します。

サンプルアドインから完全な動作マニフェストを表示するには、[この Github リポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)を参照してください。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>共同編集

Excel on the web および Office 365 サブスクリプションに接続された Windows では、Excel での coauthor が可能です。 ブックでユーザー設定の関数を使用している場合、共同編集の仕事仲間に対して、カスタム関数のアドインを読み込むように求めるメッセージが表示されます。 両方のアドインを読み込んだ後、カスタム関数は共同編集によって結果を共有します。

共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。

## <a name="known-issues"></a>既知の問題

既知の問題については、[Excel カスタム関数についての GitHub のレポート](https://github.com/OfficeDev/Excel-Custom-Functions/issues)を参照してください。

## <a name="next-steps"></a>次の手順

カスタム関数を試してみましょう。 もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。

独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。 独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。

## <a name="see-also"></a>関連項目 
* [カスタム関数の要件](custom-functions-requirement-sets.md)
* [名前付けのガイドライン](custom-functions-naming.md)
* [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](make-custom-functions-compatible-with-xll-udf.md)
