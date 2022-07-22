---
title: Office アドインでの単体テスト
description: Office JavaScript API を呼び出すコードを単体テストする方法について説明します。
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21858a68734ca5d07621f3e9c88b147ebac7dde6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958749"
---
# <a name="unit-testing-in-office-add-ins"></a>Office アドインでの単体テスト

単体テストでは、Office アプリケーションへの接続など、ネットワーク接続やサービス接続を必要とせずにアドインの機能を確認します。 [Office JavaScript API を](../develop/understanding-the-javascript-api-for-office.md)呼び出 *さない* サーバー側コードとクライアント側コードの単体テストは、Office アドインと同じであるため、特別なドキュメントは必要ありません。 ただし、Office JavaScript API を呼び出すクライアント側のコードはテストが困難です。 これらの問題を解決するために、単体テストでのモック Office オブジェクトの作成を簡略化するライブラリを作成しました。 [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock) です。 ライブラリを使用すると、次の方法でテストが簡単になります。

- Office JavaScript API は、Office アプリケーション (Excel、Word など) のコンテキストで Web ビュー コントロールで初期化する必要があります。そのため、開発用コンピューターで単体テストを実行するプロセスでは読み込めません。 Office-Addin-Mock ライブラリをテスト ファイルにインポートできます。これにより、テストを実行するnode.js プロセス内で Office JavaScript API をモックできます。
- [アプリケーション固有の API には](../develop/understanding-the-javascript-api-for-office.md#api-models)、他の関数と相互に対して特定の順序で呼び出す必要がある[読み込み](../develop/application-specific-api-model.md#load)メソッドと[同期](../develop/application-specific-api-model.md#sync)メソッドがあります。 さらに、テストする関数の`load`*後* でコードによって読み込まれる Office オブジェクトのプロパティに応じて、特定のパラメーターを使用してメソッドを呼び出す必要があります。 しかし、単体テスト フレームワークは本質的にステートレスであるため、呼び出されたかどうか `load` 、 `sync` または渡されたパラメーターの記録を `load`保持することはできません。 Office-Addin-Mock ライブラリを使用して作成するモック オブジェクトには、これらのことを追跡する内部状態があります。 これにより、モック オブジェクトは実際の Office オブジェクトのエラー動作をエミュレートできます。 たとえば、テスト対象の関数が最初に渡されなかったプロパティの読み取りを `load`試みる場合、テストは Office が返すエラーと同様のエラーを返します。

ライブラリは Office JavaScript API に依存せず、次のような JavaScript 単体テスト フレームワークで使用できます。

- [冗談](https://jestjs.io)
- [モカ](https://mochajs.org/)
- [ジャスミン](https://jasmine.github.io/)

この記事の例では、Jest フレームワークを使用します。 [Office-Addin-Mock ホーム ページ](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples)には、Mocha フレームワークを使用する例があります。

## <a name="prerequisites"></a>前提条件

この記事では、単体テストとモック作成の基本的な概念 (テスト ファイルの作成と実行方法など) について理解していることと、単体テスト フレームワークに関する経験があることを前提としています。

> [!TIP]
> Visual Studio を使用している場合は、Visual Studio での JavaScript 単体テストに関する基本的な情報については、 [Visual Studio の JavaScript と TypeScript の単体テスト](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) に関する記事を参照してから、この記事に戻ってください。

## <a name="install-the-tool"></a>ツールをインストールする

ライブラリをインストールするには、コマンド プロンプトを開き、アドイン プロジェクトのルートに移動して、次のコマンドを入力します。

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>基本的な使用法

1. プロジェクトには 1 つ以上のテスト ファイルが含まれます。 (テスト フレームワークの手順と、以下の例 (#examples) のテスト ファイルの例を参照してください)。次の例に示すように、 `require` Office JavaScript API を呼び出す関数のテストを含むテスト ファイルに、ライブラリをキーワードで `import` インポートします。

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. またはキーワードを使用してテストするアドイン関数を含むモジュールを`require``import`インポートします。 次に示すのは、テスト ファイルがアドインのコード ファイルを含むフォルダーのサブフォルダーにあることを前提とした例です。

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. 関数をテストするためにモックする必要があるプロパティとサブプロパティを持つデータ オブジェクトを作成します。 Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) プロパティと [Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) メソッドをモックするオブジェクトの例を次に示します。 これは最終的なモック オブジェクトではありません。 最終的なモック オブジェクトを作成するために使用 `OfficeMockObject` されるシード オブジェクトと考えてください。

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. データ オブジェクトをコンストラクターに `OfficeMockObject` 渡します。 返されるオブジェクトについては、次の点に `OfficeMockObject` 注意してください。

   - これは、 [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) オブジェクトの簡略化されたモックです。
   - モック オブジェクトには、データ オブジェクトのすべてのメンバーが含まれており、また、そのモックの実装 `load` と `sync` メソッドもあります。
   - モック オブジェクトは、オブジェクトの重大なエラー動作を `ClientRequestContext` 模倣します。 たとえば、テスト中の Office API が最初にプロパティを読み込んで呼び出 `sync`さずにプロパティを読み取ろうとした場合、テストは、運用環境ランタイムでスローされるエラーと同様のエラーで失敗します。"エラー、プロパティは読み込まれません" です。

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > この型の `OfficeMockObject` 完全なリファレンス ドキュメントは [、Office-Addin-Mock にあります](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)。

1. テスト フレームワークの構文で、関数のテストを追加します。 モックする `OfficeMockObject` オブジェクトの代わりにオブジェクト (この場合はオブジェクト) を `ClientRequestContext` 使用します。 Jest の例を次に示します。 このテスト例では、テスト対象のアドイン関数が呼び出 `getSelectedRangeAddress`され、オブジェクトがパラメーターとして受け取られ `ClientRequestContext` 、現在選択されている範囲のアドレスを返すことを想定しています。 完全な例については、 [この記事の後半で説明します](#mocking-a-clientrequestcontext-object)。

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. テスト フレームワークと開発ツールのドキュメントに従ってテストを実行します。 通常、テスト フレームワークを実行するスクリプトを含む **package.json** ファイルがあります。 たとえば、Jest がフレームワークの場合、 **package.json** には次のものが含まれます。

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   テストを実行するには、プロジェクトのルートにあるコマンド プロンプトに次のように入力します。

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>例

このセクションの例では、既定の設定で Jest を使用します。 これらの設定では、CommonJS モジュールがサポートされています。 Jest と node.jsを構成して ECMAScript モジュールをサポートし、TypeScript をサポートする方法については、 [Jest のドキュメント](https://jestjs.io/docs/getting-started) を参照してください。 これらの例のいずれかを実行するには、次の手順に従います。

1. 適切な Office ホスト アプリケーション (Excel や Word など) 用の Office アドイン プロジェクトを作成します。 これを迅速に行う方法の 1 つは、 [Office アドイン用の Yeoman ジェネレーターを使用する方法です](../develop/yeoman-generator-overview.md)。
1. プロジェクトのルートに [Jest をインストールします](https://jestjs.io/docs/getting-started)。
1. [office-addin-mock ツールをインストールします](#install-the-tool)。
1. 例の最初のファイルとまったく同じようにファイルを作成し、プロジェクトの他のソース ファイル (多くの場合 `\src`は .
1. ソース ファイル フォルダーにサブフォルダーを作成し、次のような `\tests`適切な名前を付けます。
1. 例のテスト ファイルとまったく同じファイルを作成し、サブフォルダーに追加します。
1. `test` **package.json** ファイルにスクリプトを追加し、[基本的な使用法](#basic-usage)の説明に従ってテストを実行します。

### <a name="mocking-the-office-common-apis"></a>Office Common API のモック

この例では、Office [Common API (](../develop/office-javascript-api-object-model.md) Excel、PowerPoint、Word など) をサポートする任意のホストの Office アドインを想定しています。 アドインには、.. という名前のファイル内の機能のいずれかが含まれています `my-common-api-add-in-feature.js`。 ファイルの内容を次に示します。 この関数は`addHelloWorldText`、"Hello World!" というテキストを設定します。 ドキュメントで現在選択されているものに対して。たとえば、Word の範囲、Excel のセル、または PowerPoint のテキスト ボックス。

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

名前付きの `my-common-api-add-in-feature.test.js` テスト ファイルは、アドイン コード ファイルの場所に対する相対サブフォルダーにあります。 ファイルの内容を次に示します。 最上位レベルのプロパティは `context`[Office.Context](/javascript/api/office/office.context) オブジェクトであるため、モックされているオブジェクトがこのプロパティの親である [Office](/javascript/api/office) オブジェクトであることに注意してください。 このコードについては、次の点に注意してください。

- コンストラクターは`OfficeMockObject`、すべての Office 列挙型クラスをモック `Office` オブジェクトに追加 *するわけではない* の`CoercionType.Text`で、アドイン メソッドで参照される値をシード オブジェクトに明示的に追加する必要があります。
- Office JavaScript ライブラリはノード プロセスに読み込まれていないため、 `Office` アドイン コードで参照されるオブジェクトを宣言して初期化する必要があります。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### <a name="mocking-the-outlook-apis"></a>Outlook API のモック

厳密に言えば、Outlook API は Common API モデルの一部ですが、 [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトを中心に構築された特別なアーキテクチャを備えています。そのため、Outlook の明確な例が提供されています。 この例では、Outlook という名前 `my-outlook-add-in-feature.js`のファイル内の機能の 1 つを持つ Outlook を想定しています。 ファイルの内容を次に示します。 この関数は`addHelloWorldText`、"Hello World!" というテキストを設定します。 メッセージ作成ウィンドウで現在選択されているものに対して設定します。

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

名前付きの `my-outlook-add-in-feature.test.js` テスト ファイルは、アドイン コード ファイルの場所に対する相対サブフォルダーにあります。 ファイルの内容を次に示します。 最上位レベルのプロパティは `context`[Office.Context](/javascript/api/office/office.context) オブジェクトであるため、モックされているオブジェクトがこのプロパティの親である [Office](/javascript/api/office) オブジェクトであることに注意してください。 このコードについては、次の点に注意してください。

- モック オブジェクトのプロパティは `host` 、Office アプリケーションを識別するためにモック ライブラリによって内部的に使用されます。 Outlook では必須です。 現在、他の Office アプリケーションでは目的を果たしません。
- Office JavaScript ライブラリはノード プロセスに読み込まれていないため、 `Office` アドイン コードで参照されるオブジェクトを宣言して初期化する必要があります。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### <a name="mocking-the-office-application-specific-apis"></a>Office アプリケーション固有の API をモックする

アプリケーション固有の API を使用する関数をテストする場合は、適切な種類のオブジェクトをモックしていることを確認してください。 次のような 2 つのオプションがあります。

- [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext) をモックします。 これは、テスト対象の関数が次の条件の両方を満たしている場合に行います。

  - *ホスト* は呼び出されません。`run` [関数 (Excel.run](/javascript/api/excel#Excel_run_batch_) など)。
  - *Host* オブジェクトの他の直接プロパティやメソッドは参照しません。

- [Excel](/javascript/api/excel) や [Word](/javascript/api/word) などの *Host* オブジェクトをモックします。 上記のオプションが不可能な場合は、この操作を行います。

両方の種類のテストの例は、以下のサブセクションにあります。

#### <a name="mocking-a-clientrequestcontext-object"></a>ClientRequestContext オブジェクトをモックする

この例では、Excel アドインの機能の 1 つがファイルに含 `my-excel-add-in-feature.js`まれているものとします。 ファイルの内容を次に示します。 これは、 `getSelectedRangeAddress` 渡 `Excel.run`されるコールバック内で呼び出されるヘルパー メソッドであることに注意してください。

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

名前付きの `my-excel-add-in-feature.test.js` テスト ファイルは、アドイン コード ファイルの場所に対する相対サブフォルダーにあります。 ファイルの内容を次に示します。 最上位レベルのプロパティは `workbook`、モックされているオブジェクトが : オブジェクトの `Excel.Workbook`親であるため、 `ClientRequestContext` 注意してください。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### <a name="mocking-a-host-object"></a>ホスト オブジェクトをモックする

この例では、Word アドインの機能の 1 つがファイルに含 `my-word-add-in-feature.js`まれているものとします。 ファイルの内容を次に示します。

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

名前付きの `my-word-add-in-feature.test.js` テスト ファイルは、アドイン コード ファイルの場所に対する相対サブフォルダーにあります。 ファイルの内容を次に示します。 最上位レベルのプロパティは `context`オブジェクト `ClientRequestContext` であるため、モックされるオブジェクトはオブジェクトというプロパティ `Word` の親になります。 このコードについては、次の点に注意してください。

- コンストラクターによって `OfficeMockObject` 最終的なモック オブジェクトが作成されると、子 `ClientRequestContext` オブジェクトに確実に含 `sync` まれるメソッドが `load` 作成されます。
- コンストラクターは`OfficeMockObject`モック `Word` オブジェクトに関数を`run`追加 *しないため*、シード オブジェクトに明示的に追加する必要があります。
- コンストラクターは`OfficeMockObject`、すべての Word 列挙型クラスをモック `Word` オブジェクトに追加 *するわけではない* の`InsertLocation.end`で、アドイン メソッドで参照される値をシード オブジェクトに明示的に追加する必要があります。
- Office JavaScript ライブラリはノード プロセスに読み込まれていないため、 `Word` アドイン コードで参照されるオブジェクトを宣言して初期化する必要があります。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run function.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

> [!NOTE]
> この型の `OfficeMockObject` 完全なリファレンス ドキュメントは [、Office-Addin-Mock にあります](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)。

## <a name="see-also"></a>関連項目

- [Office-Addin-Mock npm ページ](https://www.npmjs.com/package/office-addin-mock) のインストール ポイント。 
- オープンソースリポジトリは [Office-Addin-Mock です](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock)。
- [冗談](https://jestjs.io)
- [モカ](https://mochajs.org/)
- [ジャスミン](https://jasmine.github.io/)
