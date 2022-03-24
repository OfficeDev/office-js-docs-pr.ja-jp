---
ms.date: 11/06/2020
description: カスタム関数Excelローカライズします。
title: カスタム関数のローカライズ
ms.localizationpriority: medium
ms.openlocfilehash: 7219c838cfd5a6c827b74b5d04442280be7ebac7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744510"
---
# <a name="localize-custom-functions"></a>カスタム関数のローカライズ

アドインとカスタム関数名の両方をローカライズできます。 これを行うには、関数の JSON ファイルにローカライズされた関数名と、XML マニフェスト ファイル内のロケール情報を指定します。

>[!IMPORTANT]
> 自動生成されたメタデータはローカライズでは機能しないので、JSON ファイルを手動で更新する必要があります。 これを行う方法については、「カスタム関数の [JSON メタデータを手動で作成する」を参照してください。](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>関数名をローカライズする

カスタム関数をローカライズするには、言語ごとに新しい JSON メタデータ ファイルを作成します。 各言語 JSON ファイルで、ターゲット言語 `name` で `description` プロパティを作成します。 英語の既定のファイルは、 **functions.json という名前です**。 **functions-de.json** など、追加の JSON ファイルごとにファイル名のロケールを使用して、それらを識別します。

と`name`に`description`表示されExcelローカライズされます。 ただし、各 `id` 関数はローカライズされません。 プロパティ`id`は、関数Excel一意として識別する方法であり、設定後に変更する必要はありません。

次の JSON は、プロパティ "MULTIPLY" を使用して関数を定義 `id` する方法を示しています。 関数 `name` の and `description` プロパティはドイツ語用にローカライズされています。 各パラメーター `name` はドイツ `description` 語にもローカライズされます。

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

英語の場合、前の JSON と次の JSON を比較します。

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

言語ごとに JSON ファイルを作成した後、各 JSON メタデータ ファイルの URL を指定するロケールごとに上書き値を使用して XML マニフェスト ファイルを更新します。 次のマニフェスト XML は、(ドイツ) `en-us` の JSON ファイル URL `de-de` を上書きする既定のロケールを示しています。 **functions-de.json ファイル** には、ローカライズされたドイツ語の関数名と id が含まれています。

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

アドインをローカライズするプロセスの詳細については、「ローカライズ for Office[」を参照してください](../develop/localization.md#control-localization-from-the-manifest)。

## <a name="next-steps"></a>次の手順
カスタム関数の [名前付け規則について説明するか、](custom-functions-naming.md) エラー処理 [のベスト プラクティスを確認します](custom-functions-errors.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
