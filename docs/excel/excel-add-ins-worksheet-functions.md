---
title: Excel JavaScript API を使用して Excel の組み込みワークシート関数を呼び出す
description: ''
ms.date: 01/24/2017
ms.openlocfilehash: e709884db0bef36f1ff9a59ebf25d000f160d043
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945807"
---
# <a name="call-built-in-excel-worksheet-functions"></a>Excel の組み込みワークシート関数の呼び出し

この記事では、Excel JavaScript API を使用して、Excel の組み込みワークシート関数 (`VLOOKUP` や `SUM` など) を呼び出す方法について説明します。 また、Excel JavaScript API を使用して呼び出し可能な Excel の組み込みワークシート関数の完全な一覧も示します。

> [!NOTE]
> Excel JavaScript API を使用して Excel の*カスタム関数*を作成する方法については、「[Excel でのカスタム関数の作成](custom-functions-overview.md)」を参照してください。

## <a name="calling-a-worksheet-function"></a>ワークシート関数の呼び出し

次のコード スニペットは、ワークシート関数の呼び出し方法を示しています。`sampleFunction()` の部分はプレースホルダーであり、呼び出す関数の名前と関数が必要とする入力パラメーターに置き換えます。 ワークシート関数から返される **FunctionResult** オブジェクトの **value** プロパティには、指定した関数の結果が格納されます。 この例に示すように、**FunctionResult** の **value** プロパティは、読み込み前に `load` しておく必要があります。 この例では、関数の結果は単にコンソールに書き出されます。 

```js
var functionResult = context.workbook.functions.sampleFunction(); 
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> Excel JavaScript API を使用して呼び出し可能な関数の一覧については、この記事の「[サポートされているワークシート関数](#supported-worksheet-functions)」のセクションを参照してください。

## <a name="sample-data"></a>サンプル データ

次の画像は、各種工具の 3 か月間の販売データを格納する Excel ワークシートのテーブルを示しています。 テーブル内のそれぞれの数値は、特定の期間に特定のツールが販売された単位数を表しています。 この後の各例では、このデータに組み込みワークシート関数を適用する方法を示します。

![11 月、12 月、および 1 月のハンマー、レンチ、およびノコギリの販売データに関する Excel のスクリーンショット](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>例 1: 単一の関数

次のコード例では、前述のサンプル データに `VLOOKUP` 関数を適用して、11 月 (November) に販売したレンチ (Wrench) の数を特定しています。

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="example-2-nested-functions"></a>例 2: 入れ子になった関数

次のコード例では、前述のサンプル データに `VLOOKUP` 関数を適用して 11 月のレンチの販売数と 12 月のレンチの販売数を特定してから、その 2 か月間に販売したレンチの合計数を計算するために `SUM` 関数を適用しています。 

この例で示すように、1 つ以上の関数呼び出しが別の関数呼び出し内で入れ子にされているときには、その後で読み取ることが必要になる最終結果 (この例では、`sumOfTwoLookups`) の `load` を実行するだけで済みます。 中間結果 (この例では、それぞれの `VLOOKUP` 関数の結果) は計算され、最終結果を計算するために使用されます。

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false), 
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="supported-worksheet-functions"></a>サポートされているワークシート関数

Excel JavaScript API を使用して呼び出し可能な Excel の組み込みワークシート関数は次のとおりです。

| 関数 | 戻り値の種類 | 説明 |
|:---------------|:-------------|:-----------|
| <a href="https://support.office.com/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">ABS 関数</a> | FunctionResult | 数値の絶対値を返します。 |
| <a href="https://support.office.com/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">ACCRINT 関数</a> | FunctionResult | 定期的に利息が支払われる証券の未収利息額を返します。 |
| <a href="https://support.office.com/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">ACCRINTM 関数</a> | FunctionResult | 満期日に利息が支払われる証券の未収利息額を返します。 |
| <a href="https://support.office.com/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">ACOS 関数</a> | FunctionResult | 数値の逆余弦 (アークコサイン) を返します。 |
| <a href="https://support.office.com/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">ACOSH 関数</a> | FunctionResult | 数値の逆双曲線余弦を返します。 |
| <a href="https://support.office.com/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">ACOT 関数</a> | FunctionResult | 数値の逆余接 (アークコタンジェント) を返します。 |
| <a href="https://support.office.com/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">ACOTH 関数</a> | FunctionResult | 数値の逆双曲線余接を返します。 |
| <a href="https://support.office.com/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">AMORDEGRC 関数</a> | FunctionResult | 減価償却係数を使用して、各会計期における減価償却費を返します。 |
| <a href="https://support.office.com/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">AMORLINC 関数</a> | FunctionResult | 各会計期における減価償却費を返します。 |
| <a href="https://support.office.com/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">AND 関数</a> | FunctionResult | すべての引数が true のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">ARABIC 関数</a> | FunctionResult | ローマ数字をアラビア数字に変換します。 |
| <a href="https://support.office.com/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">AREAS 関数</a> | FunctionResult | 指定の範囲に含まれる領域の個数を返します。 |
| <a href="https://support.office.com/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">ASC 関数</a> | FunctionResult | 全角 (2 バイト) の英数カナ文字を半角 (1 バイト) の文字に変換します。 |
| <a href="https://support.office.com/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">ASIN 関数</a> | FunctionResult | 数値の逆正弦 (アークサイン) を返します。 |
| <a href="https://support.office.com/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c" target="_blank">ASINH 関数</a> | FunctionResult | 数値の逆双曲線正弦を返します。 |
| <a href="https://support.office.com/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">ATAN 関数</a> | FunctionResult | 数値の逆正接 (アークタンジェント) を返します。 |
| <a href="https://support.office.com/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">ATAN2 関数</a> | FunctionResult | 指定された x-y 座標の逆正接 (アークタンジェント) を返します。 |
| <a href="https://support.office.com/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">ATANH 関数</a> | FunctionResult | 数値の逆双曲線正接を返します。 |
| <a href="https://support.office.com/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">AVEDEV 関数</a> | FunctionResult | データ全体の平均値に対するそれぞれのデータの絶対偏差の平均を返します。 |
| <a href="https://support.office.com/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">AVERAGE 関数</a> | FunctionResult | 引数の平均値を返します。 |
| <a href="https://support.office.com/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">AVERAGEA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む引数の平均値を返します。 |
| <a href="https://support.office.com/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">AVERAGEIF 関数</a> | FunctionResult | 範囲内の検索条件に一致するすべてのセルの平均値 (算術平均) を返します。 |
| <a href="https://support.office.com/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">AVERAGEIFS 関数</a> | FunctionResult | 複数の検索条件に一致するすべてのセルの平均値 (算術平均) を返します。 |
| <a href="https://support.office.com/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">BAHTTEXT 関数</a> | FunctionResult | バーツ (ß) 通貨書式を使用して、数値を文字列に変換します。 |
| <a href="https://support.office.com/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">BASE 関数</a> | FunctionResult | 数値を、指定された基数 (底) のテキスト表現に変換します。 |
| <a href="https://support.office.com/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">BESSELI 関数</a> | FunctionResult | 修正ベッセル関数 In(x) を返します。 |
| <a href="https://support.office.com/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">BESSELJ 関数</a> | FunctionResult | ベッセル関数 Jn(x) を返します。 |
| <a href="https://support.office.com/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">BESSELK 関数</a> | FunctionResult | 修正ベッセル関数 Kn(x) を返します。 |
| <a href="https://support.office.com/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">BESSELY 関数</a> | FunctionResult | ベッセル関数 Yn(x) を返します。 |
| <a href="https://support.office.com/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">BETA.DIST 関数</a> | FunctionResult | β 分布の累積分布関数の値を返します。 |
| <a href="https://support.office.com/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">BETA.INV 関数</a> | FunctionResult | 指定された β 分布の累積分布関数の逆関数値を返します。 |
| <a href="https://support.office.com/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">BIN2DEC 関数</a> | FunctionResult | 2 進数を 10 進数に変換します。 |
| <a href="https://support.office.com/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">BIN2HEX 関数</a> | FunctionResult | 2 進数を 16 進数に変換します。 |
| <a href="https://support.office.com/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">BIN2OCT 関数</a> | FunctionResult | 2 進数を 8 進数に変換します。 |
| <a href="https://support.office.com/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">BINOM.DIST 関数</a> | FunctionResult | 二項分布の確率関数の値を返します。 |
| <a href="https://support.office.com/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595" target="_blank">BINOM.DIST.RANGE 関数</a> | FunctionResult | 二項分布を使用した試行結果の確率を返します。 |
| <a href="https://support.office.com/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">BINOM.INV 関数</a> | FunctionResult | 累積二項分布の値が基準値以下になるような最小の値を返します。 |
| <a href="https://support.office.com/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">BITAND 関数</a> | FunctionResult | 2 つの数値のビット演算 AND を返します。 |
| <a href="https://support.office.com/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">BITLSHIFT 関数</a> | FunctionResult | 左に移動数ビット (shift_amount) 移動する数値を返します。 |
| <a href="https://support.office.com/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">BITOR 関数</a> | FunctionResult | 2 つの数値のビット演算 OR を返します。 |
| <a href="https://support.office.com/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">BITRSHIFT 関数</a> | FunctionResult | 右に移動数ビット (shift_amount) 移動する数値を返します。 |
| <a href="https://support.office.com/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">BITXOR 関数</a> | FunctionResult | 2 つの数値のビット演算 "排他的 OR" を返します。 |
| <a href="https://support.office.com/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">CEILING.MATH 関数</a> | FunctionResult | 数値を最も近い整数、または基準値に最も近い倍数に切り上げます。 |
| <a href="https://support.office.com/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">CEILING.PRECISE 関数</a> | FunctionResult | 数値を最も近い整数、または基準値に最も近い倍数に切り上げます。数値の符号に関係なく、切り上げます。 |
| <a href="https://support.office.com/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">CHAR 関数</a> | FunctionResult | コード番号で指定された文字を返します。 |
| <a href="https://support.office.com/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">CHISQ.DIST 関数</a> | FunctionResult | 累積 β 確率密度関数の値を返します。 |
| <a href="https://support.office.com/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">CHISQ.DIST.RT 関数</a> | FunctionResult | カイ 2 乗分布の片側確率の値を返します。 |
| <a href="https://support.office.com/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f" target="_blank">CHISQ.INV 関数</a> | FunctionResult | 累積 β 確率密度関数の値を返します。 |
| <a href="https://support.office.com/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">CHISQ.INV.RT 関数</a> | FunctionResult | カイ 2 乗分布の片側確率の逆関数値を返します。 |
| <a href="https://support.office.com/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">CHOOSE 関数</a> | FunctionResult | 値のリストから値を選択します。 |
| <a href="https://support.office.com/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">CLEAN 関数</a> | FunctionResult | 印刷できない文字を文字列から削除します。 |
| <a href="https://support.office.com/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">CODE 関数</a> | FunctionResult | テキスト文字列内の先頭文字の数値コードを返します。 |
| <a href="https://support.office.com/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">COLUMNS 関数</a> | FunctionResult | 指定の範囲に含まれる列数を返します。 |
| <a href="https://support.office.com/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">COMBIN 関数</a> | FunctionResult | 指定された個数のオブジェクトを選択するときの組み合わせの数を返します。 |
| <a href="https://support.office.com/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">COMBINA 関数</a> | FunctionResult | 指定された個数の項目を選択するときの組み合わせ (反復あり) の数を返します |
| <a href="https://support.office.com/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">COMPLEX 関数</a> | FunctionResult | 実数係数と虚数係数を、複素数に変換します。 |
| <a href="https://support.office.com/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">CONCATENATE 関数</a> | FunctionResult | 複数の文字列を 1 つの文字列に結合します。 |
| <a href="https://support.office.com/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">CONFIDENCE.NORM 関数</a> | FunctionResult | 母集団の平均に対する信頼区間を返します。 |
| <a href="https://support.office.com/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">CONFIDENCE.T 関数</a> | FunctionResult | スチューデントの t 分布を使用して、母集団の平均に対する信頼区間を返します。 |
| <a href="https://support.office.com/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">CONVERT 関数</a> | FunctionResult | 数値の単位を変換します。 |
| <a href="https://support.office.com/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">COS 関数</a> | FunctionResult | 数値の余弦 (コサイン) を返します。 |
| <a href="https://support.office.com/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">COSH 関数</a> | FunctionResult | 数値の双曲線余弦を返します。 |
| <a href="https://support.office.com/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">COT 関数</a> | FunctionResult | 角度のコタンジェントを返します。 |
| <a href="https://support.office.com/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">COTH 関数</a> | FunctionResult | 双曲線余接 (ハイパーボリック コタンジェント) を返します。 |
| <a href="https://support.office.com/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">COUNT 関数</a> | FunctionResult | 引数リストに含まれる数値の個数をカウントします。 |
| <a href="https://support.office.com/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">COUNTA 関数</a> | FunctionResult | 引数リストに含まれる値の個数をカウントします。 |
| <a href="https://support.office.com/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">COUNTBLANK 関数</a> | FunctionResult | 指定された範囲に含まれる空白セルの個数をカウントします。 |
| <a href="https://support.office.com/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">COUNTIF 関数</a> | FunctionResult | 指定された範囲に含まれるセルのうち、検索条件に一致するセルの個数をカウントします。 |
| <a href="https://support.office.com/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">COUNTIFS 関数</a> | FunctionResult | 指定された範囲に含まれるセルのうち、複数の検索条件に一致するセルの個数を返します。 |
| <a href="https://support.office.com/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">COUPDAYBS 関数</a> | FunctionResult | 利払期間の第 1 日目から受渡日までの日数を返します。 |
| <a href="https://support.office.com/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">COUPDAYS 関数</a> | FunctionResult | 受渡日を含む利払期間内の日数を返します。 |
| <a href="https://support.office.com/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">COUPDAYSNC 関数</a> | FunctionResult | 受渡日から次の利払日までの日数を返します。 |
| <a href="https://support.office.com/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">COUPNCD 関数</a> | FunctionResult | 受渡日後の次の利払日を返します。 |
| <a href="https://support.office.com/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">COUPNUM 関数</a> | FunctionResult | 受渡日と満期日の間の利払回数を返します。 |
| <a href="https://support.office.com/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">COUPPCD 関数</a> | FunctionResult | 受渡日の直前の利払日を返します。 |
| <a href="https://support.office.com/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1" target="_blank">CSC 関数</a> | FunctionResult | 角度の余割 (コセカント) を返します。 |
| <a href="https://support.office.com/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">CSCH 関数</a> | FunctionResult | 角度の双曲線余割を返します。 |
| <a href="https://support.office.com/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">CUMIPMT 関数</a> | FunctionResult | 指定の期間に支払われる利息の累計を返します。 |
| <a href="https://support.office.com/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">CUMPRINC 関数</a> | FunctionResult | 指定期間に、貸付金に対して支払われる元金の累計を返します。 |
| <a href="https://support.office.com/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">DATE 関数</a> | FunctionResult | 指定された日付に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">DATEVALUE 関数</a> | FunctionResult | 日付を表す文字列をシリアル値に変換します。 |
| <a href="https://support.office.com/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">DAVERAGE 関数</a> | FunctionResult | 選択したデータベース レコードの平均値を返します。 |
| <a href="https://support.office.com/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">DAY 関数</a> | FunctionResult | シリアル値を日付に変換します。 |
| <a href="https://support.office.com/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df" target="_blank">DAYS 関数</a> | FunctionResult | 2 つの日付間の日数を返します。 |
| <a href="https://support.office.com/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">DAYS360 関数</a> | FunctionResult | 1 年を 360 日として、2 つの日付間の日数を返します。 |
| <a href="https://support.office.com/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">DB 関数</a> | FunctionResult | 定率法 (Fixed-declining Balance Method) を利用して、特定の期における資産の減価償却費を返します。 |
| <a href="https://support.office.com/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">DBCS 関数</a> | FunctionResult | 文字列内の半角 (1 バイト) の英数カナ文字を全角 (2 バイト) の文字に変換します。 |
| <a href="https://support.office.com/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">DCOUNT 関数</a> | FunctionResult | データベース内にある数値を含むセルの個数をカウントします。 |
| <a href="https://support.office.com/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">DCOUNTA 関数</a> | FunctionResult | データベース内にある空白でないセルの個数をカウントします。 |
| <a href="https://support.office.com/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">DDB 関数</a> | FunctionResult | 倍額定率法 (Double-declining Balance Method) または指定した他の方法を使用して、特定の期における資産の減価償却費を返します。 |
| <a href="https://support.office.com/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">DEC2BIN 関数</a> | FunctionResult | 10 進数を 2 進数に変換します。 |
| <a href="https://support.office.com/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">DEC2HEX 関数</a> | FunctionResult | 10 進数を 16 進数に変換します。 |
| <a href="https://support.office.com/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">DEC2OCT 関数</a> | FunctionResult | 10 進数を 8 進数に変換します。 |
| <a href="https://support.office.com/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e" target="_blank">DECIMAL 関数</a> | FunctionResult | 指定された底の数値のテキスト表現を 10 進数に変換します。 |
| <a href="https://support.office.com/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">DEGREES 関数</a> | FunctionResult | ラジアンを度に変換します。 |
| <a href="https://support.office.com/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">DELTA 関数</a> | FunctionResult | 2 つの値が等しいかどうかをテストします。 |
| <a href="https://support.office.com/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">DEVSQ 関数</a> | FunctionResult | 偏差の平方和を返します。 |
| <a href="https://support.office.com/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">DGET 関数</a> | FunctionResult | 指定された条件に一致する 1 つのレコードをデータベースから抽出します。 |
| <a href="https://support.office.com/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">DISC 関数</a> | FunctionResult | 証券に対する割引率を返します。 |
| <a href="https://support.office.com/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">DMAX 関数</a> | FunctionResult | 選択したデータベース レコードの最大値を返します。 |
| <a href="https://support.office.com/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">DMIN 関数</a> | FunctionResult | 選択したデータベース レコードの最小値を返します。 |
| <a href="https://support.office.com/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">DOLLAR 関数</a> | FunctionResult | ドル ($) 通貨書式を使用して、数値を文字列に変換します。 |
| <a href="https://support.office.com/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427" target="_blank">DOLLARDE 関数</a> | FunctionResult | 分数で表されたドル単位の価格を、小数表示のドル価格に変換します。 |
| <a href="https://support.office.com/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">DOLLARFR 関数</a> | FunctionResult | 小数で表されたドル単位の価格を、分数表示のドル価格に変換します。 |
| <a href="https://support.office.com/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">DPRODUCT 関数</a> | FunctionResult | データベース内の、条件に一致するレコードの特定のフィールド値を乗算します。 |
| <a href="https://support.office.com/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">DSTDEV 関数</a> | FunctionResult | 選択したデータベース レコードの標本に基づいて、標準偏差の推定値を返します。 |
| <a href="https://support.office.com/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">DSTDEVP 関数</a> | FunctionResult | 選択したデータベース レコードの母集団全体に基づいて標準偏差を算出します。 |
| <a href="https://support.office.com/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">DSUM 関数</a> | FunctionResult | データベース内の、条件に一致するレコードのフィールド列にある数値を合計します。 |
| <a href="https://support.office.com/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">DURATION 関数</a> | FunctionResult | 定期的に利子が支払われる証券の年間のマコーレー デュレーションを返します。 |
| <a href="https://support.office.com/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">DVAR 関数</a> | FunctionResult | 選択したデータベース レコードの標本に基づく分散の推定値を返します。 |
| <a href="https://support.office.com/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">DVARP 関数</a> | FunctionResult | 選択したデータベース レコードの母集団全体に基づく分散を算出します。 |
| <a href="https://support.office.com/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">EDATE 関数</a> | FunctionResult | 開始日から起算して、指定した月数だけ前または後の日付に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">EFFECT 関数</a> | FunctionResult | 実効年利率を返します。 |
| <a href="https://support.office.com/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">EOMONTH 関数</a> | FunctionResult | 指定した月数だけ前または後の月の最終日に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">ERF 関数</a> | FunctionResult | 誤差関数の値を返します。 |
| <a href="https://support.office.com/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">ERF.PRECISE 関数</a> | FunctionResult | 誤差関数の値を返します。 |
| <a href="https://support.office.com/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">ERFC 関数</a> | FunctionResult | 相補誤差関数の値を返します。 |
| <a href="https://support.office.com/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">ERFC.PRECISE 関数</a> | FunctionResult | x から無限大の範囲で、相補誤差関数の積分値を返します。 |
| <a href="https://support.office.com/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">ERROR.TYPE 関数</a> | FunctionResult | エラーの種類に対応する数値を返します。 |
| <a href="https://support.office.com/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">EVEN 関数</a> | FunctionResult | 指定された数値を最も近い偶数に切り上げた値を返します。 |
| <a href="https://support.office.com/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926" target="_blank">EXACT 関数</a> | FunctionResult | 2 つのテキスト値が等しいかどうかを判定します。 |
| <a href="https://support.office.com/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">EXP 関数</a> | FunctionResult | e を底とする数値のべき乗を返します。 |
| <a href="https://support.office.com/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">EXPON.DIST 関数</a> | FunctionResult | 指数分布を返します。 |
| <a href="https://support.office.com/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">F.DIST 関数</a> | FunctionResult | F 分布の確率関数の値を返します。 |
| <a href="https://support.office.com/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">F.DIST.RT 関数</a> | FunctionResult | F 分布の確率関数の値を返します。 |
| <a href="https://support.office.com/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">F.INV 関数</a> | FunctionResult | F 分布の確率関数の逆関数の値を返します。 |
| <a href="https://support.office.com/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">F.INV.RT 関数</a> | FunctionResult | F 分布の確率関数の逆関数の値を返します。 |
| <a href="https://support.office.com/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">FACT 関数</a> | FunctionResult | 数値の階乗を返します。 |
| <a href="https://support.office.com/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">FACTDOUBLE 関数</a> | FunctionResult | 数値の二重階乗を返します。 |
| <a href="https://support.office.com/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">FALSE 関数</a> | FunctionResult | 論理値 `FALSE` を返します。 `FALSE` |
| <a href="https://support.office.com/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">FIND 関数、FINDB 関数</a> | FunctionResult | 指定されたテキスト値を他のテキスト値の中で検索します。大文字と小文字は区別されます。 |
| <a href="https://support.office.com/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">FISHER 関数</a> | FunctionResult | フィッシャー変換の値を返します。 |
| <a href="https://support.office.com/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">FISHERINV 関数</a> | FunctionResult | フィッシャー変換の逆関数値を返します。 |
| <a href="https://support.office.com/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">FIXED 関数</a> | FunctionResult | 数値を、一定の桁数のテキストとして書式設定します。 |
| <a href="https://support.office.com/article/FLOOR-function-14bb497c-24f2-4e04-b327-b0b4de5a8886" target="_blank">FLOOR 関数</a> | FunctionResult | 数値を指定された桁数で切り捨てます。 |
| <a href="https://support.office.com/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">FLOOR.MATH 関数</a> | FunctionResult | 最も近い整数値、または基準値の倍数のうちで最も近い値に切り下げます。 |
| <a href="https://support.office.com/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">FLOOR.PRECISE 関数</a> | FunctionResult | 最も近い整数値、または基準値の倍数のうちで最も近い値に切り下げます。数値の符号に関係なく、切り下げます。 |
| <a href="https://support.office.com/article/FORECAST-function-50ca49c9-7b40-4892-94e4-7ad38bbeda99" target="_blank">FORECAST 関数</a> | FunctionResult | 線形トレンドに沿った値を返します。 |
| <a href="https://support.office.com/article/FORECASTETS-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">FORECAST.ETS 関数</a> | FunctionResult | 指数平滑化 (ETS) アルゴリズムの AAA バージョンを使って、既存の (履歴) 値に基づき将来の値を返します。 |
| <a href="https://support.office.com/article/FORECASTETSCONFINT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">FORECAST.ETS.CONFINT 関数</a> | FunctionResult | 指定した対象の日付における、予測値に対する信頼区間を返します。 |
| <a href="https://support.office.com/article/FORECASTETSSEASONALITY-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">FORECAST.ETS.SEASONALITY 関数</a> | FunctionResult | 指定した時系列に関して Excel が検出した繰り返しパターンの長さを返します。 |
| <a href="https://support.office.com/article/FORECASTETSSTAT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">FORECAST.ETS.STAT 関数</a> | FunctionResult | 時系列予測の結果として統計値を返します。 |
| <a href="https://support.office.com/article/FORECASTLINEAR-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">FORECAST.LINEAR 関数</a> | FunctionResult | 既存の値に基づいて将来値を返します。 |
| <a href="https://support.office.com/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">FV 関数</a> | FunctionResult | 投資の将来価値を返します。 |
| <a href="https://support.office.com/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">FVSCHEDULE 関数</a> | FunctionResult | 一連の金利を複利計算することにより、初期投資した元金の将来の価値を返します。 |
| <a href="https://support.office.com/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">GAMMA 関数</a> | FunctionResult | Gamma 関数値を返します。 |
| <a href="https://support.office.com/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">GAMMA.DIST 関数</a> | FunctionResult | ガンマ分布の値を返します。 |
| <a href="https://support.office.com/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">GAMMA.INV 関数</a> | FunctionResult | ガンマの累積分布の逆関数値を返します。 |
| <a href="https://support.office.com/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">GAMMALN 関数</a> | FunctionResult | ガンマ関数 Γ(x) の値の自然対数を返します。 |
| <a href="https://support.office.com/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">GAMMALN.PRECISE 関数</a> | FunctionResult | ガンマ関数 Γ(x) の値の自然対数を返します。 |
| <a href="https://support.office.com/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">GAUSS 関数</a> | FunctionResult | 標準正規分布の累積分布関数より 0.5 小さい値を返します。 |
| <a href="https://support.office.com/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">GCD 関数</a> | FunctionResult | 最大公約数を返します。 |
| <a href="https://support.office.com/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">GEOMEAN 関数</a> | FunctionResult | 相乗平均を返します。 |
| <a href="https://support.office.com/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">GESTEP 関数</a> | FunctionResult | 数値がしきい値以上であるかどうかをテストします。 |
| <a href="https://support.office.com/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">HARMEAN 関数</a> | FunctionResult | 調和平均を返します。 |
| <a href="https://support.office.com/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1" target="_blank">HEX2BIN 関数</a> | FunctionResult | 16 進数を 2 進数に変換します。 |
| <a href="https://support.office.com/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">HEX2DEC 関数</a> | FunctionResult | 16 進数を 10 進数に変換します。 |
| <a href="https://support.office.com/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">HEX2OCT 関数</a> | FunctionResult | 16 進数を 8 進数に変換します。 |
| <a href="https://support.office.com/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">HLOOKUP 関数</a> | FunctionResult | 配列の上端行で特定の値を検索し、対応するセルの値を返します。 |
| <a href="https://support.office.com/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">HOUR 関数</a> | FunctionResult | シリアル値を時刻に変換します。 |
| <a href="https://support.office.com/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">HYPERLINK 関数</a> | FunctionResult | ネットワーク サーバー、イントラネット、またはインターネット上に格納されているドキュメントを開くショートカットまたはジャンプを作成します。 |
| <a href="https://support.office.com/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">HYPGEOM.DIST 関数</a> | FunctionResult | 超幾何分布を返します。 |
| <a href="https://support.office.com/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">IF 関数</a> | FunctionResult | 実行する論理テストを指定します。 |
| <a href="https://support.office.com/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">IMABS 関数</a> | FunctionResult | 指定した複素数の絶対値を返します。 |
| <a href="https://support.office.com/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">IMAGINARY 関数</a> | FunctionResult | 指定した複素数の虚数係数を返します。 |
| <a href="https://support.office.com/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">IMARGUMENT 関数</a> | FunctionResult | 偏角シータを (ラジアンで表した角度で) 返します。 |
| <a href="https://support.office.com/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">IMCONJUGATE 関数</a> | FunctionResult | 複素数の複素共役を返します。 |
| <a href="https://support.office.com/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">IMCOS 関数</a> | FunctionResult | 複素数のコサインを返します。 |
| <a href="https://support.office.com/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">IMCOSH 関数</a> | FunctionResult | 複素数の双曲線余弦を返します。 |
| <a href="https://support.office.com/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">IMCOT 関数</a> | FunctionResult | 複素数の余接 (コタンジェント) を返します。 |
| <a href="https://support.office.com/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">IMCSC 関数</a> | FunctionResult | 複素数の余割 (コセカント) を返します。 |
| <a href="https://support.office.com/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">IMCSCH 関数</a> | FunctionResult | 複素数の双曲線余割を返します。 |
| <a href="https://support.office.com/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">IMDIV 関数</a> | FunctionResult | 2 つの複素数の商を返します。 |
| <a href="https://support.office.com/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">IMEXP 関数</a> | FunctionResult | 複素数のべき乗を返します。 |
| <a href="https://support.office.com/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">IMLN 関数</a> | FunctionResult | 複素数の自然対数を返します。 |
| <a href="https://support.office.com/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">IMLOG10 関数</a> | FunctionResult | 複素数の 10 を底とする対数 (常用対数) を返します。 |
| <a href="https://support.office.com/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">IMLOG2 関数</a> | FunctionResult | 複素数の 2 を底とする対数を返します。 |
| <a href="https://support.office.com/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">IMPOWER 関数</a> | FunctionResult | 複素数の整数乗を返します。 |
| <a href="https://support.office.com/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">IMPRODUCT 関数</a> | FunctionResult | 2 から 255 個の複素数の積を返します。 |
| <a href="https://support.office.com/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">IMREAL 関数</a> | FunctionResult | 複素数の実数係数を返します。 |
| <a href="https://support.office.com/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">IMSEC 関数</a> | FunctionResult | 複素数の正割 (セカント) を返します。 |
| <a href="https://support.office.com/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b" target="_blank">IMSECH 関数</a> | FunctionResult | 複素数の双曲線正割を返します。 |
| <a href="https://support.office.com/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">IMSIN 関数</a> | FunctionResult | 複素数の正弦を返します。 |
| <a href="https://support.office.com/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">IMSINH 関数</a> | FunctionResult | 複素数の双曲線正弦を返します。 |
| <a href="https://support.office.com/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">IMSQRT 関数</a> | FunctionResult | 複素数の平方根を返します。 |
| <a href="https://support.office.com/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">IMSUB 関数</a> | FunctionResult | 2 つの複素数の差を返します。 |
| <a href="https://support.office.com/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">IMSUM 関数</a> | FunctionResult | 複素数の和を返します。 |
| <a href="https://support.office.com/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">IMTAN 関数</a> | FunctionResult | 複素数の正接 (タンジェント) を返します。 |
| <a href="https://support.office.com/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">INT 関数</a> | FunctionResult | 指定された数値を最も近い整数に切り捨てます。 |
| <a href="https://support.office.com/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">INTRATE 関数</a> | FunctionResult | 全額投資された証券の利率を返します。 |
| <a href="https://support.office.com/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">IPMT 関数</a> | FunctionResult | 投資の指定された期に支払われる金利を返します。 |
| <a href="https://support.office.com/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">IRR 関数</a> | FunctionResult | 一連のキャッシュ フローに対する内部利益率を返します。 |
| <a href="https://support.office.com/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISERR 関数</a> | FunctionResult | 値が #N/A 以外のエラー値のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISERROR 関数</a> | FunctionResult | 値が任意のエラー値のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">ISEVEN 関数</a> | FunctionResult | 数値が偶数のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">ISFORMULA 関数</a> | FunctionResult | 数式が含まれるセルへの参照がある場合に `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISLOGICAL 関数</a> | FunctionResult | 値が論理値のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNA 関数</a> | FunctionResult | 値がエラー値 #N/A のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNONTEXT 関数</a> | FunctionResult | 値がテキスト以外のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNUMBER 関数</a> | FunctionResult | 値が数値のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">ISO.CEILING 関数</a> | FunctionResult | 最も近い整数に切り上げた値、または、指定された基準値の倍数のうち最も近い値を返します。 |
| <a href="https://support.office.com/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISODD 関数</a> | FunctionResult | 数値が奇数のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">ISOWEEKNUM 関数</a> | FunctionResult | 指定された日付のその年における ISO 週番号を返します。 |
| <a href="https://support.office.com/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">ISPMT 関数</a> | FunctionResult | 投資の指定された期に支払われる金利を計算します。 |
| <a href="https://support.office.com/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISREF 関数</a> | FunctionResult | 値が参照のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISTEXT 関数</a> | FunctionResult | 値がテキストのときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">KURT 関数</a> | FunctionResult | データ セットの尖度を返します。 |
| <a href="https://support.office.com/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">LARGE 関数</a> | FunctionResult | 指定されたデータ セットの中で k 番目に大きなデータを返します。 |
| <a href="https://support.office.com/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">LCM 関数</a> | FunctionResult | 最小公倍数を返します。 |
| <a href="https://support.office.com/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">LEFT 関数、LEFTB 関数</a> | FunctionResult | 文字列の先頭 (左端) から指定された文字数の文字を返します。 |
| <a href="https://support.office.com/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">LEN 関数、LENB 関数</a> | FunctionResult | 文字列に含まれる文字数を返します。 |
| <a href="https://support.office.com/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">LN 関数</a> | FunctionResult | 数値の自然対数を返します。 |
| <a href="https://support.office.com/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">LOG 関数</a> | FunctionResult | 指定された数を底とする数値の対数を返します。 |
| <a href="https://support.office.com/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">LOG10 関数</a> | FunctionResult | 10 を底とする数値の対数 (常用対数) を返します。 |
| <a href="https://support.office.com/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">LOGNORM.DIST 関数</a> | FunctionResult | 対数の累積分布の値を返します。 |
| <a href="https://support.office.com/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">LOGNORM.INV 関数</a> | FunctionResult | 対数の累積分布の逆関数値を返します。 |
| <a href="https://support.office.com/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb" target="_blank">LOOKUP 関数</a> | FunctionResult | ベクトルまたは配列を検索して、対応する値を返します。 |
| <a href="https://support.office.com/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">LOWER 関数</a> | FunctionResult | テキストを小文字に変換します。 |
| <a href="https://support.office.com/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">MATCH 関数</a> | FunctionResult | 参照または配列で値を検索します。 |
| <a href="https://support.office.com/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">MAX 関数</a> | FunctionResult | 引数リストに含まれる最大値を返します。 |
| <a href="https://support.office.com/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">MAXA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む引数リストから最大値を返します。 |
| <a href="https://support.office.com/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">MDURATION 関数</a> | FunctionResult | 額面価格を $100 と仮定して、証券に対する修正済マコーレー デュレーションを返します。 |
| <a href="https://support.office.com/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">MEDIAN 関数</a> | FunctionResult | 指定された数値のメジアン (中央値) を返します。 |
| <a href="https://support.office.com/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">MID 関数、MIDB 関数</a> | FunctionResult | 文字列の任意の位置から指定された文字数の文字を返します。 |
| <a href="https://support.office.com/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">MIN 関数</a> | FunctionResult | 引数リストに含まれる最小値を返します。 |
| <a href="https://support.office.com/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">MINA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む引数リストから最小値を返します。 |
| <a href="https://support.office.com/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">MINUTE 関数</a> | FunctionResult | シリアル値を時刻の分に変換します。 |
| <a href="https://support.office.com/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">MIRR 関数</a> | FunctionResult | 支払い (負の値) と収益 (正の値) のキャッシュ フローがさまざまな率で行われる場合の修正内部利益率を返します。 |
| <a href="https://support.office.com/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">MOD 関数</a> | FunctionResult | 除算の剰余を返します。 |
| <a href="https://support.office.com/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">MONTH 関数</a> | FunctionResult | シリアル値を月に変換します。 |
| <a href="https://support.office.com/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">MROUND 関数</a> | FunctionResult | 指定された値の倍数になるように、数値を四捨五入します。 |
| <a href="https://support.office.com/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">MULTINOMIAL 関数</a> | FunctionResult | 指定された複数の数値の多項係数を返します。 |
| <a href="https://support.office.com/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">N 関数</a> | FunctionResult | 値を数値に変換します。 |
| <a href="https://support.office.com/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">NA 関数</a> | FunctionResult | エラー値 #N/A を返します。 |
| <a href="https://support.office.com/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">NEGBINOM.DIST 関数</a> | FunctionResult | 負の二項分布を返します。 |
| <a href="https://support.office.com/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">NETWORKDAYS 関数</a> | FunctionResult | 2 つの日付間の稼働日の日数を返します。 |
| <a href="https://support.office.com/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">NETWORKDAYS.INTL 関数</a> | FunctionResult | 週末がどの曜日で何日間あるかを示すパラメーターを使用して、2 つの日付間にある稼働日の日数を返します。 |
| <a href="https://support.office.com/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">NOMINAL 関数</a> | FunctionResult | 名目年利率を返します。 |
| <a href="https://support.office.com/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">NORM.DIST 関数</a> | FunctionResult | 正規分布の累積分布の値を返します。 |
| <a href="https://support.office.com/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">NORM.INV 関数</a> | FunctionResult | 正規分布の累積分布の逆関数値を返します。 |
| <a href="https://support.office.com/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">NORM.S.DIST 関数</a> | FunctionResult | 標準正規分布の累積分布の値を返します。 |
| <a href="https://support.office.com/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">NORM.S.INV 関数</a> | FunctionResult | 標準正規分布の累積分布の逆関数値を返します。 |
| <a href="https://support.office.com/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">NOT 関数</a> | FunctionResult | 引数の論理値を逆にして返します。 |
| <a href="https://support.office.com/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">NOW 関数</a> | FunctionResult | 現在の日付と時刻に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">NPER 関数</a> | FunctionResult | 投資に必要な期間を返します。 |
| <a href="https://support.office.com/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">NPV 関数</a> | FunctionResult | 定期的に発生する一連のキャッシュ フローと割引率に基づいて、投資の正味現在価値を返します。 |
| <a href="https://support.office.com/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">NUMBERVALUE 関数</a> | FunctionResult | 文字列をロケールに依存しない方法で数値に変換します。 |
| <a href="https://support.office.com/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">OCT2BIN 関数</a> | FunctionResult | 8 進数を 2 進数に変換します。 |
| <a href="https://support.office.com/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">OCT2DEC 関数</a> | FunctionResult | 8 進数を 10 進数に変換します。 |
| <a href="https://support.office.com/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f" target="_blank">OCT2HEX 関数</a> | FunctionResult | 8 進数を 16 進数に変換します。 |
| <a href="https://support.office.com/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">ODD 関数</a> | FunctionResult | 指定された数値を最も近い奇数に切り上げた値を返します。 |
| <a href="https://support.office.com/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">ODDFPRICE 関数</a> | FunctionResult | 1 期目の日数が半端な証券に対して、額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">ODDFYIELD 関数</a> | FunctionResult | 1 期目の日数が半端な証券の利回りを返します。 |
| <a href="https://support.office.com/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">ODDLPRICE 関数</a> | FunctionResult | 最終期の日数が半端な証券に対して、額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">ODDLYIELD 関数</a> | FunctionResult | 最終期の日数が半端な証券の利回りを返します。 |
| <a href="https://support.office.com/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">OR 関数</a> | FunctionResult | いずれかの引数が true のときに `TRUE` を返します。 |
| <a href="https://support.office.com/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf" target="_blank">PDURATION 関数</a> | FunctionResult | 投資が指定した価値に達するまでの投資期間を返します。 |
| <a href="https://support.office.com/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">PERCENTILE.EXC 関数</a> | FunctionResult | 特定の範囲に含まれるデータの第 k 百分位数に当たる値を返します (k は 0 より大きく 1 より小さい値)。 |
| <a href="https://support.office.com/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">PERCENTILE.INC 関数</a> | FunctionResult | 特定の範囲に含まれるデータの第 k 百分位数に当たる値を返します。 |
| <a href="https://support.office.com/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">PERCENTRANK.EXC 関数</a> | FunctionResult | データ セット内での値の順位を百分率 (0 より大きく 1 より小さい) で表した値を返します。 |
| <a href="https://support.office.com/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">PERCENTRANK.INC 関数</a> | FunctionResult | データ セット内での値の順位を百分率で表した値を返します。 |
| <a href="https://support.office.com/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">PERMUT 関数</a> | FunctionResult | 指定された個数のオブジェクトを選択するときの順列の数を返します。 |
| <a href="https://support.office.com/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">PERMUTATIONA 関数</a> | FunctionResult | すべてのオブジェクトから指定された数のオブジェクト (繰り返しを含む) を選択する場合の順列の数を返します。 |
| <a href="https://support.office.com/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">PHI 関数</a> | FunctionResult | 標準正規分布の密度関数の値を返します。 |
| <a href="https://support.office.com/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">PI 関数</a> | FunctionResult | 円周率 π を返します。 |
| <a href="https://support.office.com/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441" target="_blank">PMT 関数</a> | FunctionResult | 年間の定期支払額を算出します。 |
| <a href="https://support.office.com/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">POISSON.DIST 関数</a> | FunctionResult | ポワソン分布の値を返します。 |
| <a href="https://support.office.com/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">POWER 関数</a> | FunctionResult | 数値のべき乗を返します。 |
| <a href="https://support.office.com/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">PPMT 関数</a> | FunctionResult | 指定した期に支払われる投資元金を返します。 |
| <a href="https://support.office.com/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">PRICE 関数</a> | FunctionResult | 定期的に利息が支払われる証券に対して、額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">PRICEDISC 関数</a> | FunctionResult | 割引証券の額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">PRICEMAT 関数</a> | FunctionResult | 満期日に利息が支払われる証券に対して、額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/PROB-function-9ac30561-c81c-4259-8253-34f0a238fc49" target="_blank">PROB 関数</a> | FunctionResult | 指定した範囲に含まれる値が上限と下限との間に収まる確率を返します。 |
| <a href="https://support.office.com/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">PRODUCT 関数</a> | FunctionResult | 引数を乗算します。 |
| <a href="https://support.office.com/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">PROPER 関数</a> | FunctionResult | 文字列に含まれる英単語の先頭文字だけを大文字に変換します。 |
| <a href="https://support.office.com/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">PV 関数</a> | FunctionResult | 投資の現在価値を返します。 |
| <a href="https://support.office.com/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">QUARTILE.EXC 関数</a> | FunctionResult | 0 より大きく 1 より小さい百分位値に基づいて、データ セットに含まれるデータから四分位数を返します。 |
| <a href="https://support.office.com/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">QUARTILE.INC 関数</a> | FunctionResult | データ セットの四分位数を返します。 |
| <a href="https://support.office.com/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">QUOTIENT 関数</a> | FunctionResult | 除算の商の整数部を返します。 |
| <a href="https://support.office.com/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">RADIANS 関数</a> | FunctionResult | 度をラジアンに変換します。 |
| <a href="https://support.office.com/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">RAND 関数</a> | FunctionResult | 0 から 1 の乱数を返します。 |
| <a href="https://support.office.com/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">RANDBETWEEN 関数</a> | FunctionResult | 指定された範囲内の数値の乱数を返します。 |
| <a href="https://support.office.com/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">RANK.AVG 関数</a> | FunctionResult | 数値のリストの中で、指定した数値の順位を返します。 |
| <a href="https://support.office.com/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40" target="_blank">RANK.EQ 関数</a> | FunctionResult | 数値のリストの中で、指定した数値の順位を返します。 |
| <a href="https://support.office.com/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">RATE 関数</a> | FunctionResult | 年間の投資金利を返します。 |
| <a href="https://support.office.com/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">RECEIVED 関数</a> | FunctionResult | 全額投資された証券に対して、満期日に支払われる金額を返します。 |
| <a href="https://support.office.com/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a" target="_blank">REPLACE 関数、REPLACEB 関数</a> | FunctionResult | テキスト内の文字を置き換えます。 |
| <a href="https://support.office.com/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">REPT 関数</a> | FunctionResult | テキストを指定した回数だけ繰り返します。 |
| <a href="https://support.office.com/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">RIGHT 関数、RIGHTB 関数</a> | FunctionResult | 文字列の末尾 (右端) から指定された文字数の文字を返します。 |
| <a href="https://support.office.com/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">ROMAN 関数</a> | FunctionResult | アラビア数字を、ローマ数字を表す文字列に変換します。 |
| <a href="https://support.office.com/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">ROUND 関数</a> | FunctionResult | 数値を四捨五入して指定された桁数にします。 |
| <a href="https://support.office.com/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">ROUNDDOWN 関数</a> | FunctionResult | 数値を指定された桁数で切り捨てます。 |
| <a href="https://support.office.com/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">ROUNDUP 関数</a> | FunctionResult | 数値を指定された桁数で切り上げます。 |
| <a href="https://support.office.com/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">ROWS 関数</a> | FunctionResult | 指定の範囲に含まれる行数を返します。 |
| <a href="https://support.office.com/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">RRI 関数</a> | FunctionResult | 投資の成長に対する等価利率を返します。 |
| <a href="https://support.office.com/article/RTD-function-e0cc001a-56f0-470a-9b19-9455dc0eb593" target="_blank">RTD 関数</a> | FunctionResult | COM オートメーションに対応するプログラムからリアルタイムのデータを取得します。 |
| <a href="https://support.office.com/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">SEC 関数</a> | FunctionResult | 角度の正割 (セカント) を返します。 |
| <a href="https://support.office.com/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">SECH 関数</a> | FunctionResult | 角度の双曲線正割を返します。 |
| <a href="https://support.office.com/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">SECOND 関数</a> | FunctionResult | シリアル値を秒に変換します。 |
| <a href="https://support.office.com/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">SERIESSUM 関数</a> | FunctionResult | 数式で定義されるべき級数の和を返します。 |
| <a href="https://support.office.com/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">SHEET 関数</a> | FunctionResult | 参照先のシートのシート番号を返します。 |
| <a href="https://support.office.com/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">SHEETS 関数</a> | FunctionResult | 参照内のシート数を返します |
| <a href="https://support.office.com/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">SIGN 関数</a> | FunctionResult | 数値の符号を返します。 |
| <a href="https://support.office.com/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">SIN 関数</a> | FunctionResult | 指定された角度のサインを返します。 |
| <a href="https://support.office.com/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">SINH 関数</a> | FunctionResult | 数値の双曲線正弦を返します。 |
| <a href="https://support.office.com/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">SKEW 関数</a> | FunctionResult | 分布の歪度を返します。 |
| <a href="https://support.office.com/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">SKEW.P 関数</a> | FunctionResult | 母集団に基づく分布の歪度を取得します。歪度とは、分布の平均値周辺での両側の非対称度を表す値です。 |
| <a href="https://support.office.com/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">SLN 関数</a> | FunctionResult | 定額法 (Straight-line Method) を使用して、資産の 1 期あたりの減価償却費を返します。 |
| <a href="https://support.office.com/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07" target="_blank">SMALL 関数</a> | FunctionResult | 指定されたデータ セットの中で k 番目に小さなデータを返します。 |
| <a href="https://support.office.com/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">SQRT 関数</a> | FunctionResult | 正の平方根を返します。 |
| <a href="https://support.office.com/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">SQRTPI 関数</a> | FunctionResult | (数値 * π) の平方根を返します。 |
| <a href="https://support.office.com/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">STANDARDIZE 関数</a> | FunctionResult | 正規化された値を返します。 |
| <a href="https://support.office.com/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">STDEV.P 関数</a> | FunctionResult | 母集団全体に基づいて、標準偏差を計算します。 |
| <a href="https://support.office.com/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">STDEV.S 関数</a> | FunctionResult | 標本に基づく標準偏差の推定値を返します。 |
| <a href="https://support.office.com/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">STDEVA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む標本に基づいて、標準偏差の推定値を返します。 |
| <a href="https://support.office.com/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">STDEVPA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む母集団全体に基づいて、標準偏差を計算します。 |
| <a href="https://support.office.com/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">SUBSTITUTE 関数</a> | FunctionResult | 文字列中の指定された文字を他の新しい文字に置き換えます。 |
| <a href="https://support.office.com/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939" target="_blank">SUBTOTAL 関数</a> | FunctionResult | リストまたはデータベースの集計値を返します。 |
| <a href="https://support.office.com/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">SUM 関数</a> | FunctionResult | 引数を合計します。 |
| <a href="https://support.office.com/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b" target="_blank">SUMIF 関数</a> | FunctionResult | 指定された検索条件に一致するセルの値を合計します。 |
| <a href="https://support.office.com/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">SUMIFS 関数</a> | FunctionResult | セル範囲内で、複数の検索条件を満たすセルの値を合計します。 |
| <a href="https://support.office.com/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">SUMSQ 関数</a> | FunctionResult | 引数の 2 乗の和 (平方和) を返します。 |
| <a href="https://support.office.com/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">SYD 関数</a> | FunctionResult | 級数法 (Sum-of-Year's Digits Method) を使用して、特定の期における減価償却費を返します。 |
| <a href="https://support.office.com/article/T-function-fb83aeec-45e7-4924-af95-53e073541228" target="_blank">T 関数</a> | FunctionResult | 引数をテキストに変換します。 |
| <a href="https://support.office.com/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">T.DIST 関数</a> | FunctionResult | スチューデントの t 分布のパーセンテージ (確率) を返します。 |
| <a href="https://support.office.com/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">T.DIST.2T 関数</a> | FunctionResult | スチューデントの t 分布のパーセンテージ (確率) を返します。 |
| <a href="https://support.office.com/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">T.DIST.RT 関数</a> | FunctionResult | スチューデントの t 分布の値を返します。 |
| <a href="https://support.office.com/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">T.INV 関数</a> | FunctionResult | スチューデントの t 分布の t 値を、確率の関数と自由度で返します。 |
| <a href="https://support.office.com/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">T.INV.2T 関数</a> | FunctionResult | スチューデントの t 分布の逆関数値を返します。 |
| <a href="https://support.office.com/article/TAN-function-08851a40-179f-4052-b789-d7f699447401" target="_blank">TAN 関数</a> | FunctionResult | 数値の正接 (タンジェント) を返します。 |
| <a href="https://support.office.com/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">TANH 関数</a> | FunctionResult | 数値の双曲線正接を返します。 |
| <a href="https://support.office.com/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">TBILLEQ 関数</a> | FunctionResult | 米国財務省短期証券 (TB) の債券換算利回りを返します。 |
| <a href="https://support.office.com/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">TBILLPRICE 関数</a> | FunctionResult | 米国財務省短期証券 (TB) の額面 $100 あたりの価格を返します。 |
| <a href="https://support.office.com/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">TBILLYIELD 関数</a> | FunctionResult | 米国財務省短期証券 (TB) の利回りを返します。 |
| <a href="https://support.office.com/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">TEXT 関数</a> | FunctionResult | 数値を、書式設定したテキストに変換します。 |
| <a href="https://support.office.com/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">TIME 関数</a> | FunctionResult | 指定した時刻に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">TIMEVALUE 関数</a> | FunctionResult | 時刻を表す文字列をシリアル値に変換します。 |
| <a href="https://support.office.com/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">TODAY 関数</a> | FunctionResult | 現在の日付に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">TRIM 関数</a> | FunctionResult | テキストからスペースを削除します。 |
| <a href="https://support.office.com/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">TRIMMEAN 関数</a> | FunctionResult | データ セットの中間項の平均を返します。 |
| <a href="https://support.office.com/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">TRUE 関数</a> | FunctionResult | 論理値 `TRUE` を返します。 `TRUE` |
| <a href="https://support.office.com/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">TRUNC 関数</a> | FunctionResult | 数値の小数部を切り捨てて整数にします。 |
| <a href="https://support.office.com/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">TYPE 関数</a> | FunctionResult | 値のデータ型を表す数値を返します。 |
| <a href="https://support.office.com/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">UNICHAR 関数</a> | FunctionResult | 指定された数値により参照される Unicode 文字を返します。 |
| <a href="https://support.office.com/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">UNICODE 関数</a> | FunctionResult | 文字列の最初の文字に対応する番号 (コード ポイント) を返します。 |
| <a href="https://support.office.com/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">UPPER 関数</a> | FunctionResult | 文字列に含まれる英字をすべて大文字に変換します。 |
| <a href="https://support.office.com/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">VALUE 関数</a> | FunctionResult | テキスト引数を数値に変換します。 |
| <a href="https://support.office.com/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">VAR.P 関数</a> | FunctionResult | 母集団全体に基づいて、分散を計算します。 |
| <a href="https://support.office.com/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b" target="_blank">VAR.S 関数</a> | FunctionResult | 標本に基づいて、分散の推定値を返します。 |
| <a href="https://support.office.com/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">VARA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む標本に基づいて、分散の推定値を返します。 |
| <a href="https://support.office.com/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">VARPA 関数</a> | FunctionResult | 数値、文字列、および論理値を含む母集団全体に基づいて、分散を計算します。 |
| <a href="https://support.office.com/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">VDB 関数</a> | FunctionResult | 定率法 (declining Balance Method) を利用して、特定の期または部分的な期における資産の減価償却費を返します。 |
| <a href="https://support.office.com/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">VLOOKUP 関数</a> | FunctionResult | 配列の左端列で特定の値を検索し、その行内で移動して、対応するセルの値を返します。 |
| <a href="https://support.office.com/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">WEEKDAY 関数</a> | FunctionResult | シリアル値を曜日に変換します。 |
| <a href="https://support.office.com/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">WEEKNUM 関数</a> | FunctionResult | シリアル値をその年の何週目に当たるかを示す値に変換します。 |
| <a href="https://support.office.com/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">WEIBULL.DIST 関数</a> | FunctionResult | ワイブル分布の値を返します。 |
| <a href="https://support.office.com/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">WORKDAY 関数</a> | FunctionResult | 指定した稼動日数だけ前または後の日付に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">WORKDAY.INTL 関数</a> | FunctionResult | 週末がどの曜日で何日間あるかを示すパラメーターを使用して、指定した稼働日数だけ前または後の日付に対応するシリアル値を返します。 |
| <a href="https://support.office.com/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">XIRR 関数</a> | FunctionResult | 定期的でないキャッシュ フローの特定のスケジュールに対する内部利益率を返します。 |
| <a href="https://support.office.com/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">XNPV 関数</a> | FunctionResult | 定期的でないキャッシュ フローの特定のスケジュールに対する正味現在価値を返します。 |
| <a href="https://support.office.com/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">XOR 関数</a> | FunctionResult | すべての引数の論理排他 OR を返します。 |
| <a href="https://support.office.com/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">YEAR 関数</a> | FunctionResult | シリアル値を年に変換します。 |
| <a href="https://support.office.com/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">YEARFRAC 関数</a> | FunctionResult | 開始日と終了日を指定して、その間の期間が 1 年間に対して占める割合を返します。 |
| <a href="https://support.office.com/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">YIELD 関数</a> | FunctionResult | 利息が定期的に支払われる証券の利回りを返します。 |
| <a href="https://support.office.com/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">YIELDDISC 関数</a> | FunctionResult | 米国財務省短期証券 (TB) などの割引債の年利回りを返します。 |
| <a href="https://support.office.com/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">YIELDMAT 関数</a> | FunctionResult | 満期日に利息が支払われる証券の利回りを返します。 |
| <a href="https://support.office.com/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Z.TEST 関数</a> | FunctionResult | Z 検定の片側確率の値を返します。 |

## <a name="see-also"></a>関連項目

- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API オープン仕様](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [ワークシート関数のオブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js)
