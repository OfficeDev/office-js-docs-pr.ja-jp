---
title: Visual Studio 2019 で JavaScript IntelliSense を利用できるようにする
description: ''
ms.date: 10/11/2019
localization_priority: Priority
ms.openlocfilehash: afe2890015c853c2b70a00d94ea1b6a49296a14a
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626783"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a><span data-ttu-id="3f6c6-102">Visual Studio 2019 で JavaScript IntelliSense を利用できるようにする</span><span class="sxs-lookup"><span data-stu-id="3f6c6-102">Get JavaScript IntelliSense in Visual Studio</span></span>

<span data-ttu-id="3f6c6-p101">Visual Studio 2019 を使用して Office アドインを開発する場合は、JSDoc を使用することで、JavaScript の変数、オブジェクト、パラメーター、および戻り値の IntelliSense を有効にできます。この記事では、JSDoc の概要と、JSDoc を使用して Visual Studio の IntellSense を作成する方法について説明します。詳細については、「[JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense)」および「[JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3f6c6-p101">When you use Visual Studio 2017 to develop Office Add-ins, you can use JSDoc to enable IntelliSense for your JavaScript variables, objects, parameters, and return values. This article provides an overview of JSDoc and how you can use it to create IntellSense in Visual Studio. For more details, see [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) and [JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span></span> 

## <a name="officejs-type-definitions"></a><span data-ttu-id="3f6c6-106">Office.js の型定義</span><span class="sxs-lookup"><span data-stu-id="3f6c6-106">Office.js type definitions</span></span>

<span data-ttu-id="3f6c6-p102">Visual Studio に Office.js の型の定義を提供する必要があります。そのために、次の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="3f6c6-p102">You need to provide the definitions of the types in Office.js to Visual Studio. To do this, you can:</span></span>

- <span data-ttu-id="3f6c6-p103">`\Office\1\` という名前のソリューションのフォルダーに Office.js ファイルのローカル コピーを用意します。アドイン プロジェクトの作成時に、Visual Studio の Office アドイン プロジェクト テンプレートにより、このローカル コピーが追加されます。</span><span class="sxs-lookup"><span data-stu-id="3f6c6-p103">Have a local copy of the Office.js files in a folder in your solution named `\Office\1\`. The Office Add-in project templates in Visual Studio add this local copy when you create an add-in project.</span></span> 
- <span data-ttu-id="3f6c6-p104">アドイン ソリューションの Web アプリケーション プロジェクトのルートに、tsconfig.json ファイルを追加することで、Office.js のオンライン バージョンを使用します。ファイルには、次のコンテンツが含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="3f6c6-p104">Use an online version of Office.js by adding a tsconfig.json file to the root of the web application project in the add-in solution. The file should include the following content.</span></span>

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

## <a name="jsdoc-syntax"></a><span data-ttu-id="3f6c6-113">JSDoc 構文</span><span class="sxs-lookup"><span data-stu-id="3f6c6-113">JSDoc syntax</span></span>

<span data-ttu-id="3f6c6-p105">基本的な手法として、変数 (またはパラメーターなど) の前に、データ型を識別するコメントを付けます。これにより、Visual Studio の IntelliSense は、そのメンバーを推測できるようになります。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="3f6c6-p105">The basic technique is to precede the variable (or parameter, and so on) with a comment that identifies its data type. This allows IntelliSense in Visual Studio to infer its members. The following are examples.</span></span>

### <a name="variable"></a><span data-ttu-id="3f6c6-117">可変</span><span class="sxs-lookup"><span data-stu-id="3f6c6-117">Variable</span></span>

```js
/** @type {Excel.Range} */
var subsetRange;
```
![変数の IntelliSense](../images/intellisense-vs17-var.png)

### <a name="parameter"></a><span data-ttu-id="3f6c6-119">パラメーター</span><span class="sxs-lookup"><span data-stu-id="3f6c6-119">Parameter</span></span>

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```
![パラメーターの IntelliSense](../images/intellisense-vs17-param.png)

### <a name="return-value"></a><span data-ttu-id="3f6c6-121">戻り値</span><span class="sxs-lookup"><span data-stu-id="3f6c6-121">Return value</span></span>

```js
/** @returns {Word.Range} */
function myFunc() {

}
```
![戻り値の IntelliSense](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a><span data-ttu-id="3f6c6-123">複合型</span><span class="sxs-lookup"><span data-stu-id="3f6c6-123">Complex types</span></span>

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```
![複合型の IntelliSense](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a><span data-ttu-id="3f6c6-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="3f6c6-125">See also</span></span>

- [<span data-ttu-id="3f6c6-126">Visual Studio でのアドインの作成とデバッグ</span><span class="sxs-lookup"><span data-stu-id="3f6c6-126">Create and debug add-ins in Visual Studio</span></span>](create-and-debug-office-add-ins-in-visual-studio.md)
