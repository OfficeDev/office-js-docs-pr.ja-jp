---
ms.date: 04/29/2020
description: Excel カスタム関数をローカライズします。
title: カスタム関数をローカライズする
localization_priority: Normal
ms.openlocfilehash: 427bff029c5e85caa216f628df450525ee187c17
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609297"
---
# <a name="localize-custom-functions"></a>カスタム関数をローカライズする

アドインとカスタム関数名の両方をローカライズできます。 そのためには、ローカライズされた関数名を、XML マニフェストファイルの関数の JSON ファイルとロケール情報に提供します。

>[!IMPORTANT]
> 自動生成されたメタデータはローカライズには機能しないため、JSON ファイルを手動で更新する必要があります。 これを行う方法については、「 [Excel のカスタム関数のメタデータ](custom-functions-json.md)」を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>関数名をローカライズする

カスタム関数をローカライズするには、言語ごとに新しい JSON メタデータファイルを作成します。 各言語 JSON ファイルで、 `name` `description` ターゲット言語でプロパティを作成します。 英語の既定のファイルの名前は、**関数 json**です。 各 JSON ファイルのファイル名にロケールを使用します。たとえば、**関数**を識別するために使用します。

は `name` 、 `description` Excel に表示され、ローカライズされています。 ただし、 `id` 各関数のはローカライズされていません。 `id`このプロパティでは、Excel によって関数が一意であると識別されますが、設定された後に変更することはできません。

次の JSON は、"掛け算" というプロパティを持つ関数を定義する方法を示して `id` います。 `name` `description` 関数のおよびプロパティは、ドイツ語にローカライズされています。 各パラメーター `name` と `description` は、ドイツ語にローカライズされています。

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

前の JSON を次の JSON と比較して英語を比較します。

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a>アドインをローカライズする

各言語の JSON ファイルを作成した後、各 JSON メタデータファイルの URL を指定する各ロケールの上書き値で XML マニフェストファイルを更新します。 次のマニフェスト XML は、 `en-us` (ドイツ) 用の JSON ファイルの上書き URL を含む既定のロケールを示して `de-de` います。 **関数の de**ファイルには、ローカライズされたドイツ語の関数名と id が含まれています。

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

アドインのローカライズプロセスの詳細については、「 [Office アドインのローカライズ](../develop/localization.md#control-localization-from-the-manifest)」を参照してください。

## <a name="next-steps"></a>次の手順
[カスタム関数の名前付け規則](custom-functions-naming.md)について、または[エラー処理のベストプラクティス](custom-functions-errors.md)を検出する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
