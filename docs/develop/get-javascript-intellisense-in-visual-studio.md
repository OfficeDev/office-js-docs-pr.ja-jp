---
title: Visual Studio で JavaScript IntelliSense を利用できるようにする
description: JSDoc を使用して、JavaScript 変数、オブジェクト、パラメーター、および戻り値の IntelliSense を作成する方法について説明します。
ms.date: 04/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fb18199c53b7b836079754b5fb4561b76a968b1
ms.sourcegitcommit: 3c5ede9c4f9782947cea07646764f76156504ff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/06/2022
ms.locfileid: "64682260"
---
# <a name="get-javascript-intellisense-in-visual-studio"></a>Visual Studio で JavaScript IntelliSense を利用できるようにする

Visual Studio 2019 以降を使用してアドインOffice開発する場合、JSDoc を使用して JavaScript 変数、オブジェクト、パラメーター、および戻り値に対して IntelliSense を有効にすることができます。 この記事では、JSDoc の概要と、JSDoc を使用して Visual Studio の IntellSense を作成する方法について説明します。 詳細については、「[JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense)」および「[JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)」を参照してください。 

## <a name="officejs-type-definitions"></a>Office.js の型定義

Visual Studio に Office.js の型の定義を提供する必要があります。そのために、次の操作を実行します。

- `\Office\1\` という名前のソリューションのフォルダーに Office.js ファイルのローカル コピーを用意します。アドイン プロジェクトの作成時に、Visual Studio の Office アドイン プロジェクト テンプレートにより、このローカル コピーが追加されます。 
- アドイン ソリューションの Web アプリケーション プロジェクトのルートに、tsconfig.json ファイルを追加することで、Office.js のオンライン バージョンを使用します。ファイルには、次のコンテンツが含まれている必要があります。

    ```json
        {
            "compilerOptions": {
                "allowJs": true,            // These settings apply to JavaScript files also.
                "noEmit":  true             // Do not compile the JS (or TS) files in this project.
            },
            "exclude": [
                "node_modules",             // Don't include any JavaScript found under "node_modules".
                "Scripts/Office/1"          // Suppress loading all the JavaScript files from the Office NuGet package.
            ],
            "typeAcquisition": {
                "enable": true,             // Enable automatic fetching of type definitions for detected JavaScript libraries.
                "include": [ "office-js" ]  // Ensure that the "Office-js" type definition is fetched.
            }
        }
    ```

## <a name="jsdoc-syntax"></a>JSDoc 構文

基本的な手法として、変数 (またはパラメーターなど) の前に、データ型を識別するコメントを付けます。これにより、Visual Studio の IntelliSense は、そのメンバーを推測できるようになります。次に例を示します。

### <a name="variable"></a>可変

```js
/** @type {Excel.Range} */
var subsetRange;
```

!['subsetRange' 変数の IntelliSense の抜粋を示すスクリーンショット。](../images/intellisense-vs17-var.png)

### <a name="parameter"></a>パラメーター

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

!['paras' パラメーター (JavaScript の例では'paragraphs' パラメーター) の IntelliSense の抜粋を示すスクリーンショット。](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>戻り値

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

!['myFunc()' 戻り値の IntelliSense の抜粋を示すスクリーンショット。](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>複合型

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![たとえば、'var myVar;' の複合型宣言の IntelliSense を示すスクリーンショット。](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a>関連項目

- [Visual Studio を使用して Office アドインを開発する](develop-add-ins-visual-studio.md)
- [Visual Studio で Office アドインをデバッグする](debug-office-add-ins-in-visual-studio.md)
