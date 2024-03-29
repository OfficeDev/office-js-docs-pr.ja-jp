---
description: Office アドインの Excel カスタム関数を作成します。
title: Excel でカスタム関数を作成する
ms.date: 08/04/2021
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 12740615215913b0201426f929dbcb803c866648
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422762"
---
# <a name="create-custom-functions-in-excel"></a>Excel でカスタム関数を作成する

開発者は、カスタム関数を使用して関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。 ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次のアニメーション画像は、JavaScript または Typescript で作成した関数を呼び出すブックを示しています。 この例では、カスタム関数 `=MYFUNCTION.SPHEREVOLUME` は球の体積を計算します。

![MYFUNCTION.SPHEREVOLUME カスタム関数を Excel ワークシートのセルへ挿入するエンド ユーザーを示すアニメーション画像。](../images/SphereVolumeNew.gif)

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

> [!TIP]
> カスタム関数アドインで作業ウィンドウまたはリボン ボタンを使用する場合は、カスタム関数コードの実行に加えて、 [共有ランタイム](../testing/runtimes.md#shared-runtime)を設定する必要があります。 詳細については、「 [共有ランタイムを使用するように Office アドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="how-a-custom-function-is-defined-in-code"></a>コードでカスタム関数を定義する方法

[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用して Excel のカスタム関数アドイン プロジェクトを作成する場合、使用する関数および作業ウィンドウを制御するファイルが作成されます。このため、カスタム関数に重要なファイルに注意を集中できます。

| ファイル | ファイル形式 | 説明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>または<br/>**./src/functions/functions.ts** | JavaScript<br/>または<br/>TypeScript | カスタム関数を定義するコードが含みます。 |
| **./src/functions/functions.html** | HTML | カスタム関数を定義する JavaScript ファイルに &lt;script&gt; 参照を提供します。 |
| **./manifest.xml** | XML | カスタム関数 JavaScript、JSON、HTML ファイルなど、カスタム関数が使用する複数のファイルの場所を指定します。 また、作業ウィンドウ ファイルおよびコマンド ファイルの場所を表示すると共に、カスタム関数が使用するランタイムも指定します。 |

### <a name="script-file"></a>スクリプト ファイル

スクリプト ファイル (**./src/customfunctions.js** または **/src/customfunctions.ts**) には、カスタム関数を定義するコードと関数を定義するコメントが含まれています。

`add` カスタム関数は次のコードにより定義されます。 コード コメントは、Excel にカスタム関数を記述する JSON メタデータ ファイルを生成するために使用されます。 必須の `@customfunction` コメントが最初に宣言されて、これがカスタム関数であることを示します。 次に、`description` プロパティに続いて、`first` および `second` の 2 つのパラメーターが宣言されます。 最後に `returns` の説明が記述されます。 カスタム関数で必要になるコメントに関する詳細については、「[カスタム関数の JSON メタデータを自動作成する](custom-functions-json-autogeneration.md)」を参照してください。

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

カスタム関数 ([Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)によって作成されたプロジェクトの **./manifest.xml**) を定義するアドイン用 XML マニフェスト ファイルには、以下のような複数の機能があります。

- カスタム関数に名前空間を定義します。ユーザーがアドインの一部として関数を特定するのに役立つように、名前空間がカスタム関数の前に付加されます。
- カスタム関数マニフェストに固有の **\<ExtensionPoint\>** および **\<Resources\>** 要素を使用します。 これらの要素には、JavaScript、JSON、および HTML ファイルの場所に関する情報が含まれています。
- カスタム関数に使用するランタイムを指定します。別のランタイムを特段必要とする場合を除いて、共有ランタイムは関数と作業ウィンドウの間でデータを共有できるため、共有ランタイムを常に使用することをお勧めします。

[Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してファイルを作成する場合は、これらのファイルの既定値ではないため、共有ランタイムを使用するようにマニフェストを調整することをお勧めします。 マニフェストを変更するには、「 [共有ランタイムを使用するように Excel アドインを構成する」の手順に](../develop/configure-your-add-in-to-use-a-shared-runtime.md)従います。

サンプル アドインからフル機能マニフェストを表示するには、[Office アドインのサンプルの Github リポジトリのアドイン サンプルのいずれか](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/excel-shared-runtime-global-state/manifest.xml)でマニフェストを参照してください。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>共同編集

Excel on the web および Microsoft 365 サブスクリプションに接続されている Windows 版の Excel では、エンド ユーザーは Excel で共同編集を行うことができます。 エンド ユーザーのブックでカスタム関数を使用している場合、そのエンド ユーザーの共同編集の仕事仲間は、対応するカスタム関数のアドインを読み込むように求められます。 両方のユーザーがアドインを読み込むと、共同編集によりカスタム関数は結果を共有します。

共同編集の詳細については、「[Excel での共同編集](/office/vba/excel/concepts/about-coauthoring-in-excel)」を参照してください。

## <a name="next-steps"></a>次の手順

カスタム関数を試してみましょう。もしまだであれば、簡単な[カスタム関数クイックスタート](../quickstarts/excel-custom-functions-quickstart.md)または、詳細な[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)をご覧ください。

独自のカスタム関数を試すもう 1 つの簡単な方法は[スクリプト ラボ](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)を使用し、アドインで Excel のカスタム関数を試してみることができます。 独自のカスタム関数を作成したり、提供されたサンプルを再生してみることができます。

## <a name="see-also"></a>関連項目

- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)
- [カスタム関数の要件セット](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
- [カスタム関数の名前付けのガイドライン](custom-functions-naming.md)
- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](make-custom-functions-compatible-with-xll-udf.md)
- [共有ランタイムを使用するように Office アドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Office アドインのランタイム](../testing/runtimes.md)
