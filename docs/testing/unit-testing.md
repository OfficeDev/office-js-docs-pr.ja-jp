---
title: アドインでの単体Officeテスト
description: JavaScript API を呼び出すテスト コードを単体Officeする方法について説明します。
ms.date: 11/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8824b8e759e3c1acecf30683f2b89bb41bd558f3
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242041"
---
# <a name="unit-testing-in-office-add-ins"></a>アドインでの単体Officeテスト

単体テストでは、ネットワーク接続やサービス接続を必要とせずに、アドインの機能を確認します (アプリケーションへの接続Officeします。 [Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md)を呼び出していないサーバー側コードとクライアント側コードの単体テストは、Office アドインの場合と Web アプリケーションの場合と同じなので、特別なドキュメントは必要としません。 ただし、JavaScript API を呼び出すクライアント側Officeテストは困難です。 これらの問題を解決するために、単体テストでのモック Office オブジェクトの作成を簡略化するためのライブラリを作成しました[。Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). ライブラリを使用すると、次の方法でテストが容易になります。

- Office JavaScript API は、Office アプリケーション (Excel、Word など) のコンテキストで webview コントロールで初期化する必要があります。そのため、開発コンピューターで単体テストを実行するプロセスに読み込む必要があります。 Office-Addin-Mock ライブラリをテスト ファイルにインポートすると、テストを実行する node.js プロセス内で Office JavaScript API をモックできます。
- アプリケーション[固有の API には](../develop/understanding-the-javascript-api-for-office.md#api-models)、他[](../develop/application-specific-api-model.md#sync)の関数と互いに対して特定の順序で呼び出す必要がある読み込みメソッドと同期メソッドがあります。 [](../develop/application-specific-api-model.md#load) さらに、テスト対象の関数で後でコードで読み取る Office オブジェクトのプロパティに応じて、メソッドを特定のパラメーターで呼び出す `load` 必要があります。  ただし、単体テスト フレームワークは本質的にステートレスなので、呼び出されたかどうか、またはどのパラメーターに渡されたのかを記録 `load` `sync` することはできません `load` 。 Office-Addin-Mock ライブラリを使用して作成するモック オブジェクトには、これらのことを追跡する内部状態があります。 これにより、モック オブジェクトは実際のオブジェクトのエラー動作Officeできます。 たとえば、テスト中の関数が、最初に渡されていないプロパティを読み取ろうとすると、テストは、Office に返されるエラーと同様のエラー `load` を返します。

ライブラリは JavaScript API のOffice依存し、次のような JavaScript 単体テスト フレームワークで使用できます。

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [ジャスミン](https://jasmine.github.io/)

この記事の例では、Jest フレームワークを使用します。 Mocha フレームワークを使用する例は[、Office-Addin-Mock のホーム ページに表示されます](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples)。

## <a name="prerequisites"></a>前提条件

この記事では、テスト ファイルの作成と実行方法など、単体テストとモックの基本的な概念に精通し、単体テスト フレームワークの経験を持っている必要があります。

> [!TIP]
> Visual Studio を使用している場合は、Visual Studio での JavaScript 単体テストに関する基本的な情報については、Visual Studio の「JavaScript と[TypeScript](/visualstudio/javascript/unit-testing-javascript-with-visual-studio)の単体テスト」の記事を読んでから、この記事に戻することをお勧めします。

## <a name="install-the-tool"></a>ツールのインストール

ライブラリをインストールするには、コマンド プロンプトを開き、アドイン プロジェクトのルートに移動し、次のコマンドを入力します。

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>基本的な使用法

1. プロジェクトには 1 つ以上のテスト ファイルがあります。 (以下の Examples(#examples) のテスト フレームワークの手順とテスト ファイルの例を参照してください。次の例に示すように、or キーワードを使用して、Office JavaScript API を呼び出す関数のテストを含むテスト ファイルにライブラリを `require` `import` インポートします。

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. or キーワードを使用してテストするアドイン関数を含むモジュールを `require` インポート `import` します。 次に、テスト ファイルがアドインのコード ファイルを含むフォルダーのサブフォルダーにあると仮定する例を示します。

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. 関数をテストするためにモックする必要があるプロパティとサブプロパティを持つデータ オブジェクトを作成します。 [Workbook.range.address](/javascript/api/excel/excel.range#address)プロパティと[Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#getSelectedRange__) Excelをモックするオブジェクトの例を次に示します。 これは最終的なモック オブジェクトではない。 最終的なモック オブジェクトを作成するために使用されるシード オブジェクト `OfficeMockObject` と考えて下さい。

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

1. データ オブジェクトをコンストラクターに渡 `OfficeMockObject` します。 返されるオブジェクトについて次の点に注意 `OfficeMockObject` してください。

   - これは [、OfficeExtension.ClientRequestContext オブジェクトの簡略化されたモック](/javascript/api/office/officeextension.clientrequestcontext) です。
   - モック オブジェクトには、データ オブジェクトのすべてのメンバーが含まれており、and メソッドのモック `load` 実装 `sync` も持っています。
   - モック オブジェクトは、オブジェクトの重大なエラー動作を模倣 `ClientRequestContext` します。 たとえば、テスト中の Office API が、最初にプロパティを読み込んで呼び出さずにプロパティの読み取りを試みる場合、テストは失敗し、実稼働ランタイムでスローされるエラーと同様のエラーが発生します `sync` 。"Error, property not loaded"。

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > この型の完全なリファレンス `OfficeMockObject` ドキュメントは[、Office-Addin-Mock にある](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)です。

1. テスト フレームワークの構文で、関数のテストを追加します。 オブジェクトを `OfficeMockObject` モックするオブジェクトの代り、この場合はオブジェクトを使用 `ClientRequestContext` します。 次の例は Jest で続きます。 このテスト例では、テスト対象のアドイン関数が呼び出され、オブジェクトをパラメーターとして受け取り、現在選択されている範囲のアドレスを返す目的で使用することを前提と `getSelectedRangeAddress` `ClientRequestContext` します。 完全な例は、 [この記事の後半です](#mocking-a-clientrequestcontext-object)。

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. テスト フレームワークと開発ツールのドキュメントに従ってテストを実行します。 通常、テスト フレームワークを実行するスクリプトを含む **package.json** ファイルがあります。 たとえば、Jest がフレームワークの場合 **、package.json** には次の情報が含まれます。

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   テストを実行するには、プロジェクトのルートにあるコマンド プロンプトに次を入力します。

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>例

このセクションの例では、既定の設定で Jest を使用します。 これらの設定は、CommonJS モジュールをサポートします。 Jest および node.js ECMAScript モジュールをサポートし、TypeScript をサポートする方法については [、Jest](https://jestjs.io/docs/getting-started) のドキュメントを参照してください。 これらの例を実行するには、次の手順を実行します。

1. 適切なOfficeホスト アプリケーション (たとえば、Officeまたは Word) 用のExcel作成します。 これを迅速に行う方法の 1 つは、Yo ツールを使用[Officeです](https://github.com/OfficeDev/generator-office)。
1. プロジェクトのルートに [Jest をインストールします](https://jestjs.io/docs/getting-started)。
1. [office-addin-mock ツールをインストールします](#install-the-tool)。
1. 例の最初のファイルとまったく同じファイルを作成し、プロジェクトの他のソース ファイル (よく呼ばれる) を含むフォルダーに追加します `\src` 。
1. ソース ファイル フォルダーにサブフォルダーを作成し、適切な名前を指定します `\tests` 。
1. 例のテスト ファイルとまったく同じファイルを作成し、サブフォルダーに追加します。
1. `test`Package.json ファイルに **スクリプトを** 追加し、「基本使用法」の説明に従ってテスト [を実行します](#basic-usage)。

### <a name="mocking-the-office-common-apis"></a>共通 API のOfficeする

この例では、Office 共通 API (Office、PowerPoint、Word など[](../develop/office-javascript-api-object-model.md)) をサポートする任意のホストの Excel アドインを想定しています。 アドインには、という名前のファイルの機能の 1 つがあります `my-common-api-add-in-feature.js` 。 ファイルの内容を次に示します。 この `addHelloWorldText` 関数は、テキスト "Hello World! ドキュメントで現在選択されているもの。たとえば、次の例を示します。Word の範囲、またはセル内のセルExcelテキスト ボックスを指定PowerPoint。

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

名前の付いたテスト ファイルは、アドイン コード ファイルの場所を基準としてサブフォルダー `my-common-api-add-in-feature.test.js` に格納されます。 ファイルの内容を次に示します。 トップ レベルのプロパティは、次 `context` の値[Office。Context](/javascript/api/office/office.context)オブジェクトなので、モックされているオブジェクトは、このプロパティの親であるオブジェクト(オブジェクトOffice[します。](/javascript/api/office) このコードについては、次の点に注意してください。

- コンストラクターは、すべての Office 列挙クラスをモック オブジェクトに追加する必要はありません。そのため、アドイン メソッドで参照される値をシード オブジェクトに明示的に追加する `OfficeMockObject`  `Office` `CoercionType.Text` 必要があります。
- JavaScript Officeはノード プロセスに読み込まれないので、アドイン コードで参照されるオブジェクトを宣言して初期化 `Office` する必要があります。

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

### <a name="mocking-the-outlook-apis"></a>API のOutlookする

厳密に言えば、Outlook API は共通 API モデルの一部ですが[、Mailbox](/javascript/api/outlook/office.mailbox)オブジェクトを中心に構築された特別なアーキテクチャを備え、Outlook の明確な例を示しました。 この例では、ファイルOutlook機能の 1 つを持つオブジェクトを想定しています `my-outlook-add-in-feature.js` 。 ファイルの内容を次に示します。 この `addHelloWorldText` 関数は、テキスト "Hello World! を、メッセージ作成ウィンドウで現在選択されているものに設定します。

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

名前の付いたテスト ファイルは、アドイン コード ファイルの場所を基準としてサブフォルダー `my-outlook-add-in-feature.test.js` に格納されます。 ファイルの内容を次に示します。 トップ レベルのプロパティは、次 `context` の値[Office。Context](/javascript/api/office/office.context)オブジェクトなので、モックされているオブジェクトは、このプロパティの親であるオブジェクト(オブジェクトOffice[します。](/javascript/api/office) このコードについては、次の点に注意してください。

- JavaScript Officeはノード プロセスに読み込まれないので、アドイン コードで参照されるオブジェクトを宣言して初期化 `Office` する必要があります。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
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

### <a name="mocking-the-office-application-specific-apis"></a>アプリケーション固有の API Officeをモックする

アプリケーション固有の API を使用する関数をテストする場合は、適切な種類のオブジェクトをモック化してください。 次のような 2 つのオプションがあります。

- [OfficeExtension.ClientRequestObject をモックします](/javascript/api/office/officeextension.clientrequestcontext)。 テスト中の関数が次の両方の条件を満たす場合は、この操作を実行します。

  - ホストを呼び出 *す必要があります*。`run` メソッド[(Excel.run](/javascript/api/excel#Excel_run_batch_)など)
  - Host オブジェクトの他の直接プロパティやメソッドは *参照* しない。

- ホスト オブジェクト *(ファイル* 名や Word など [) をExcel](/javascript/api/excel)[します](/javascript/api/word)。 前のオプションが使用できない場合は、この操作を行います。

両方の種類のテストの例を以下のサブセクションに示します。

#### <a name="mocking-a-clientrequestcontext-object"></a>ClientRequestContext オブジェクトのモック

この例では、ファイルExcel機能の 1 つを持つ、新しいアドインを想定しています `my-excel-add-in-feature.js` 。 ファイルの内容を次に示します。 に渡 `getSelectedRangeAddress` されるコールバック内で呼び出されるヘルパー メソッドである点に注意してください `Excel.run` 。

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

名前の付いたテスト ファイルは、アドイン コード ファイルの場所を基準としてサブフォルダー `my-excel-add-in-feature.test.js` に格納されます。 ファイルの内容を次に示します。 トップ レベルのプロパティは、モックされているオブジェクトが: オブジェクトの親である点に `workbook` `Excel.Workbook` 注意 `ClientRequestContext` してください。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectRange method.
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

#### <a name="mocking-a-host-object"></a>ホスト オブジェクトのモック

この例では、 という名前のファイルに 1 つの機能を持つ Word アドインを想定しています `my-word-add-in-feature.js` 。 ファイルの内容を次に示します。

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

名前の付いたテスト ファイルは、アドイン コード ファイルの場所を基準としてサブフォルダー `my-word-add-in-feature.test.js` に格納されます。 ファイルの内容を次に示します。 トップ レベル のプロパティはオブジェクトなので、モックされているオブジェクトは、このプロパティの親であるオブジェクト `context` `ClientRequestContext` `Word` です。 このコードについては、次の点に注意してください。

- コンストラクターが `OfficeMockObject` 最終的なモック オブジェクトを作成すると、子オブジェクトと `ClientRequestContext` メソッドが確実 `sync` に `load` 作成されます。
- コンストラクターはモック オブジェクトにメソッドを追加しないので、シード オブジェクトに明示的に追加 `OfficeMockObject`  `run` `Word` する必要があります。
- コンストラクターは、すべての Word 列挙クラスをモック オブジェクトに追加する必要はありません。そのため、アドイン メソッドで参照される値をシード オブジェクトに明示的に追加 `OfficeMockObject`  `Word` `InsertLocation.end` する必要があります。
- JavaScript Officeはノード プロセスに読み込まれないので、アドイン コードで参照されるオブジェクトを宣言して初期化 `Word` する必要があります。

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
  // Mock the Word.run method.
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

## <a name="adding-mock-objects-properties-and-methods-dynamically-when-testing"></a>テスト時にモック オブジェクト、プロパティ、およびメソッドを動的に追加する

一部のシナリオでは、効率的なテストでは、実行時にモック オブジェクトを作成または変更する必要があります。つまり、テストの実行中です。 次に、例を示します。

- テスト中の関数は、2 回目の呼び出し時の動作が異なります。 最初に 1 つのモック オブジェクトで関数をテストしてから、このモック オブジェクトを変更し、変更されたモック オブジェクトで関数を再度テストする必要があります。
- 複数の類似したが同一ではないモック オブジェクトに対して関数をテストする必要があります。 たとえば、color プロパティを持つモック オブジェクトを使用して関数をテストし、テキスト プロパティを持つモック オブジェクトを使用して関数を再度テストする必要がありますが、それ以外の場合は元のモック オブジェクトと同じです。

これらの `OfficeMockObject` シナリオで役立つ 3 つの方法があります。

- `OfficeMockObject.setMock` オブジェクトにプロパティと値を追加 `OfficeMockObject` します。 次の使用例は、プロパティを追加 `address` します。

    ```javascript
    rangeMock.setMock("address", "G6:K9");
    ```

- `OfficeMockObject.addMockFunction` 次の例に示すように `OfficeMockObject` 、オブジェクトにモック関数を追加します。

    ```javascript
    workbookMock.addMockFunction("getSelectedRange", function () { 
      const range = {
        address: "B2:G5",
      };
      return range;
    });
    ```

    > [!NOTE]
    > function パラメーターは省略可能です。 存在しない場合は、空の関数が作成されます。

- `OfficeMockObject.addMock` 新しいオブジェクト `OfficeMockObject` をプロパティとして既存のオブジェクトに追加し、名前を付けます。 これは、すべてのメンバーが持つ最小メンバー (など) `OfficeMockObject` `load` を持つ `sync` 必要があります。 and メソッドを使用して、追加のメンバー `setMock` を `addMockFunction` 追加できます。 次に、モック オブジェクトをプロパティとしてモック ブックに追加 `Excel.WorkbookProtection` `protection` する例を示します。 次に、新 `protected` しいモック オブジェクトにプロパティを追加します。

    ```javascript
    workbookMock.addMock("protection");
    workbookMock.protection.setMock("protected", true);
    ```

> [!NOTE]
> この型の完全なリファレンス `OfficeMockObject` ドキュメントは[、Office-Addin-Mock にある](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)です。

## <a name="see-also"></a>関連項目

- [Office-Addin-Mock npm ページのインストール](https://www.npmjs.com/package/office-addin-mock)ポイント。 
- オープンソースの repo は[Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock)です。
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [ジャスミン](https://jasmine.github.io/)
