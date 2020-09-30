---
ms.date: 09/25/2020
description: JSDoc タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Normal
ms.openlocfilehash: 151dc7c97b2a98743906b7e0a920fdc1eff62e7f
ms.sourcegitcommit: 42202d7e2ac24dffa77cf937f5697a1cd79ee790
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2020
ms.locfileid: "48308538"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>カスタム関数用の JSON メタデータの自動生成

Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、[JSDoc タグ](https://jsdoc.app/)が使用されます。 JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。 JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。

関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。 詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。

### <a name="adding-a-description-to-a-function"></a>関数に説明を追加する

説明は、カスタム関数の機能を理解するためのヘルプが必要な場合に、ヘルプ テキストとしてユーザーに表示されます。 説明に特定のタグは必要ありません。 JSDoc コメントに簡単な説明を入力するだけです。 一般に、説明は JSDoc コメント セクションの先頭に配置されますが、配置場所に関係なく機能します。

組み込み関数の説明の例を表示するには、Excel を開き、**[数式]** タブに移動し、**[関数の​​挿入]** を選択します。 すべての関数の説明を参照したり、独自のカスタム関数を一覧表示したりすることができます。

次の例では、「球の体積を計算します。」 が、カスタム関数の説明です。

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a>JSDoc タグ

Excel カスタム関数では、次の JSDoc タグがサポートされています。

* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id 名
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ 名前の説明
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a>@cancelable

関数が取り消されたときにカスタム関数が処理を実行することを示します。

最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。 関数は、関数 `oncanceled` が取り消されたときの結果を示すために、関数をプロパティに割り当てることができます。

最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。

関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。

---
<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

構文: @customfunction _id_ _名_

このタグは、JavaScript/TypeScript 関数が Excel カスタム関数であることを示します。 カスタム関数のメタデータを作成する必要があります。

このタグの例を次に示します。

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

は、 `id` カスタム関数を識別します。

* `id` が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。
* `id` はすべてのカスタム関数で一意である必要があります。
* 指定できる文字は、A から Z、a から z、0 から 9、アンダースコア (\_)、ピリオド (.) に制限されます。

次の例では、インクリメントは関数の `id` と `name` です。

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>name

カスタム関数の表示用の `name` を提供します。

* name が指定されていない場合、id が名前としても使用されます。
* 使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。
* 最初の文字は、アルファベット文字にする必要があります。
* 最大文字数は 128 文字です。

次の例では、INC は関数の`id` で、 `increment` は`name`です。

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>説明

Excel のユーザーには、関数の入力時に説明が表示され、関数の動作を指定します。 説明に特定のタグは必要ありません。 JSDoc コメント内に関数の機能を説明するフレーズを入力して、カスタム関数に説明を追加します。 既定では、JSDoc コメント セクションでタグが付けられていないテキストは、関数の説明です。

次の例では、「2 つの数値を加算する関数」というフレーズが、ID プロパティ `ADD` のカスタム関数の説明です。

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
<a id="helpurl"></a>

### <a name="helpurl"></a>@helpurl

構文: @helpurl _url_

指定された _url_ が Excel で表示されます。

次の例では、 `helpurl` がに `www.contoso.com/weatherhelp` なります。

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
<a id="param"></a>

### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

JavaScript 構文: @param {type} 名_の説明_

* `{type}` 中かっこで囲まれた型情報を指定します。 使用できる型に関する詳細については、「[型](#types)」セクションを参照してください。 種類が指定されていない場合は、既定の種類が使用され `any` ます。
* `name` @param タグを適用するパラメーターを指定します。 これは必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 省略可能です。

カスタム関数内のパラメーターを省略可能と指定する方法:

* パラメーター名を角かっこで囲みます。 例: `@param {string} [text] Optional text`。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

次の例は、2つまたは3つの数字を省略可能なパラメーターとして追加する ADD 関数を示しています。

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

TypeScript 構文: @param 名 _の説明_

* `name` @param タグを適用するパラメーターを指定します。 これは必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 省略可能です。

使用できる関数のパラメーターの型に関する詳細については、「[型](#types)」セクションを参照してください。

カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。

* 省略可能なパラメーターを使用する。 例: `function f(text?: string)`
* パラメーターに既定値を指定する。 例: `function f(text: string = "abc")`

@param の詳しい説明については、「[JSDoc](https://jsdoc.app/tags-param.html)」を参照してください。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

次の例は、2 つの数値を加算する `add` 関数を示しています。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requiresAddress

関数が評価されているセルのアドレスを指定する必要があることを示します。

最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。 関数が呼び出されると、`address` プロパティにアドレスが含まれます。

---
<a id="returns"></a>

### <a name="returns"></a>@returns

構文: @returns {_type_}

戻り値の型を指定します。

`{type}` を省略すると、TypeScript の型情報が使用されます。 型情報がない場合、型は `any` になります。

次の例は、 `@returns` タグを使用する `add` 関数を示しています。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
<a id="streaming"></a>

### <a name="streaming"></a>@streaming

カスタム関数がストリーミング関数であることを示すのに使用されます。 

最後のパラメーターの型は `CustomFunctions.StreamingInvocation<ResultType>` です。
関数が戻り `void` ます。

ストリーミング関数は、値を直接返すのではなく、 `setResult(result: ResultType)` 最後のパラメーターを使用して呼び出します。

ストリーム関数によってスローされる例外は無視されます。 `setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。 ストリーミング関数と詳細については、「[ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。

ストリーミング関数は、[@volatile](#volatile) としてマークできません。

---
<a id="volatile"></a>

### <a name="volatile"></a>@volatile

揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる関数です。 Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。 このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。

ストリーミング関数に揮発性関数は使用できません。

次の関数は揮発性で、 `@volatile` タグを使用します。

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>型

パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。 型が`any`の場合、変換は実行されません。

### <a name="value-types"></a>値の型

1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。

### <a name="matrix-type"></a>マトリックス型

2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。 たとえば、`number[][]`の型は数字のマトリックスを示します。 `string[][]` は、文字列のマトリックスを示します。

### <a name="error-type"></a>エラーの種類

非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。

ストリーミング関数は、エラーの種類で `setResult()` を返してエラーを示すことができます。

### <a name="promise"></a>Promise

関数は Promise を返すことができます。これは、promise が解決されたときに値を提供します。 Promise が拒否されると、エラーがスローされます。

### <a name="other-types"></a>その他の型

その他の型は、エラーとして処理されます。

## <a name="next-steps"></a>次の手順

[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。 または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
