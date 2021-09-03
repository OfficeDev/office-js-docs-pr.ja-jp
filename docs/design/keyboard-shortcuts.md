---
title: カスタム キーボード ショートカット (Office アドイン)
description: カスタム キーボード ショートカット (キーの組み合わせとも呼ばれる) をアドインに追加するOffice説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 2ac9a83511fc29eb055ebdc4d2c77f7675c68994
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868408"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>カスタム キーボード ショートカットをアドインOffice追加する

キーボード ショートカット (キーの組み合わせとも呼ばれる) を使用すると、アドインのユーザーの作業効率が向上します。 キーボード ショートカットは、マウスの代替手段を提供することで、障がいを持つユーザーに対するアドインのアクセシビリティも向上します。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> キーボード ショートカットが既に有効になっているアドインの作業バージョンから開始するには、キーボード ショートカットのサンプルを複製Excel[実行します](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。 キーボード ショートカットを独自のアドインに追加する準備ができたら、この記事に進む。

アドインにキーボード ショートカットを追加するには、3 つの手順があります。

1. [アドインのマニフェストを構成します](#configure-the-manifest)。
1. [アクションとそのキーボード ショートカットを](#create-or-edit-the-shortcuts-json-file) 定義するショートカット JSON ファイルを作成または編集します。
1. [](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API の 1 つ以上のランタイム呼び出しを追加して、各アクションに関数をマップします。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストには 2 つの小さな変更があります。 1 つは、共有ランタイムを使用するアドインを有効にし、もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルをポイントすることです。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム キーボード ショートカットを追加するには、共有ランタイムを使用するアドインが必要です。 詳細については、「 [共有ランタイムを使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>マッピング ファイルをマニフェストにリンクする

マニフェスト *内* の要素の直下 (内部ではない) `<VersionOverrides>` に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。 後の手順で作成するプロジェクトの JSON ファイルの完全な URL に属性 `Url` を設定します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>ショートカット JSON ファイルを作成または編集する

プロジェクトに JSON ファイルを作成します。 ファイルのパスが ExtendedOverrides 要素の属性に指定した場所と一致 `Url` [する必要](../reference/manifest/extendedoverrides.md) があります。 このファイルには、キーボード ショートカットと、キーボード ショートカットが呼び出すアクションが記述されます。

1. JSON ファイル内には、2 つの配列があります。 actions 配列には、呼び出すアクションを定義するオブジェクトが含まれます。ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。 次に例を示します。
    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    JSON オブジェクトの詳細については、「アクション オブジェクトを作成 [する」および](#construct-the-action-objects) 「ショートカット オブジェクトを作成 [する」を参照してください](#construct-the-shortcut-objects)。 ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

    > [!NOTE]
    > この記事では、"Ctrl" の代りで "CONTROL" を使用できます。

    後の手順では、アクション自体が作成する関数にマップされます。 この例では、後で SHOWTASKPANE をメソッドを呼び出す関数にマップし、HIDETASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップ `Office.addin.hide` します。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>アクションの関数へのマッピングを作成する

1. プロジェクトで、HTML ページによって読み込まれた JavaScript ファイルを要素で開 `<FunctionFile>` きます。
1. JavaScript ファイルで[、Office.actions.associate](/javascript/api/office/office.actions#associate) API を使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。 次の JavaScript をファイルに追加します。 コードについて次の点に注意してください。

    - 最初のパラメーターは、JSON ファイルからのアクションの 1 つです。
    - 2 番目のパラメーターは、JSON ファイル内のアクションにマップされているキーの組み合わせをユーザーが押すと実行される関数です。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. この例を続行するには、最初 `'SHOWTASKPANE'` のパラメーターとして使用します。
1. 関数の本文では[、Office.addin.showTaskpane](/javascript/api/office/office.addin#showAsTaskpane__)メソッドを使用してアドインの作業ウィンドウを開きます。 完了したら、コードは次のようになります。

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. 2 番目の関数呼び出しを追加して、アクション `Office.actions.associate` `HIDETASKPANE` を[Office.addin.hide を](/javascript/api/office/office.addin#hide__)呼び出す関数にマップします。 次に例を示します。

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

前の手順に従うと **、Ctrl** + Alt + Up キーと Ctrl + Alt + Down キーを押して、作業ウィンドウの表示を切り替 **えます**。 同じ動作は、Excelアドイン[](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)PnP repo の Officeキーボード ショートカット のサンプルにGitHub。

## <a name="details-and-restrictions"></a>詳細と制限

### <a name="construct-the-action-objects"></a>アクション オブジェクトを作成する

次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`actions` します。

- プロパティ名 `id` と `name` 必須です。
- この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。
- プロパティ `name` は、アクションを記述するユーザーフレンドリーな文字列である必要があります。 文字 A - Z、a - z、0 ~ 9、および句読点 "-"、"_"、および "+" の組み合わせである必要があります。
- プロパティは省略可能です。 現在は `ExecuteFunction` 型のみサポートされています。

次に例を示します。

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

### <a name="construct-the-shortcut-objects"></a>ショートカット オブジェクトを作成する

次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`shortcuts` します。

- プロパティ名 `action` 、 `key` および `default` 必須です。
- プロパティの値は `action` 文字列であり、action オブジェクトのプロパティの 1 `id` つと一致する必要があります。
- プロパティ `default` には、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、"+" の任意の組み合わせを指定できます。 (慣例では、これらのプロパティでは小文字は使用されません)。
- プロパティ `default` には、少なくとも 1 つの修飾子キー (Alt、Ctrl、Shift) の名前と、他の 1 つのキーのみを含む必要があります。
- Shift を唯一の修飾子キーとして使用することはできません。 Shift キーと Alt キーまたは Ctrl キーを組み合わせます。
- Mac では、Command 修飾子キーもサポートしています。
- Mac の場合、Alt は Option キーにマップされます。 このWindows、Command は Ctrl キーにマップされます。
- 標準キーボードで 2 つの文字が同じ物理キーにリンクされている場合は、プロパティ内の類義語になります。たとえば、Alt+a と Alt+A は同じショートカットなので `default` 、"-" と "_" は同じ物理キーなので、Ctrl + + と Ctrl+ も同じです。 \_
- "+" 文字は、そのいずれかの側のキーが同時に押された状態を示します。

次に例を示します。

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!NOTE]
> キーヒント (Excel ショートカットなどのシーケンシャル キー ショートカットとも呼ばれる) は、Office アドインではサポートされていません。

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>他のアドインで使用されるキーの組み合わせを回避する

ユーザーが既に使用しているキーボード ショートカットは多数Office。 既に使用されているアドインのキーボード ショートカットを登録しないようにしますが、既存のキーボード ショートカットを上書きしたり、同じキーボード ショートカットを登録した複数のアドイン間の競合を処理する必要がある場合があります。

競合が発生した場合、ユーザーが最初に競合するキーボード ショートカットを使用しようとすると、ダイアログ ボックスが表示されます。このダイアログに表示されるアクション名は、ファイル内のアクション オブジェクトのプロパティです。 `name` `shortcuts.json`

![1 つのショートカットに対して 2 つの異なるアクションを持つ競合モーダルを示す図。](../images/add-in-shortcut-conflict-modal.png)

ユーザーは、キーボード ショートカットで実行する操作を選択できます。 選択を行った後、同じショートカットの今後の使用のために基本設定が保存されます。 ショートカットの基本設定は、プラットフォームごとにユーザーごとに保存されます。 ユーザーが自分の設定を変更する場合は、[教えて]検索ボックスから [Office アドインのショートカット設定のリセット] コマンド **を** 呼び出します。 このコマンドを呼び出すと、ユーザーのすべてのアドイン ショートカット設定がクリアされ、次に競合するショートカットを使用しようとすると、ユーザーに競合ダイアログ ボックスが表示されます。

![[アドインのショートカットの設定] Excel設定のリセットOfficeを表示するダイアログ ボックスを表示します。](../images/add-in-reset-shortcuts-action.png)

最適なユーザー エクスペリエンスを得る場合は、これらの優れたプラクティスを使用して、Excelを最小限にすることをお勧めします。

- キーボード ショートカットのみを使用して、次のパターンを使用します。 **Ctrl + Shift + Alt +* x***、x は他のキーです。 
- さらにキーボード ショートカットが必要な場合は、[](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)キーボード ショートカットExcel一覧を確認し、アドインで使用しないようにします。
- キーボード フォーカスがアドイン UI 内にある場合 **、Ctrl + Spacebar** と **Ctrl + Shift + F10** は基本的なアクセシビリティ ショートカットとして機能しません。
- Windows または Mac コンピューターで、検索メニューで [Office アドインのショートカット設定をリセットする] コマンドが使用できない場合は、コンテキスト メニューからリボンをカスタマイズしてリボンにコマンドを手動で追加できます。

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>プラットフォームごとにキーボード ショートカットをカスタマイズする

ショートカットをプラットフォーム固有にカスタマイズできます。 次に、次の各プラットフォームのショートカットをカスタマイズするオブジェクトの例を `shortcuts` 示します。 `windows` `mac` `web` ただし、ショートカットごとにショートカット キー `default` が必要です。

次の例では、 `default` キーは、指定されていないプラットフォームのフォールバック キーです。 指定されていない唯一のプラットフォームはWindows、キーはユーザーにのみ `default` 適用Windows。

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a>キーボード ショートカット JSON をローカライズする

アドインが複数のローカライズをサポートしている場合は、アクション オブジェクトのプロパティをローカライズ `name` する必要があります。 また、アドインがサポートするローカライズの中にアルファベットや異なる書き込みシステムがある場合、キーボードが異なる場合は、ショートカットのローカライズも必要な場合があります。 キーボード ショートカット JSON をローカライズする方法については、「拡張オーバーライドをローカライズする [」を参照してください](../develop/localization.md#localize-extended-overrides)。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>オーバーライドできないブラウザー のショートカット

Web でカスタム キーボード ショートカットを使用する場合、ブラウザーで使用される一部のキーボード ショートカットをアドインで上書きすることはできません。このリストは進行中の作業です。 上書きできない他の組み合わせを発見した場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="next-steps"></a>次の手順

- キーボード ショートカット[Excelアドインの例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を参照してください。
- 「マニフェストの拡張オーバーライドを処理する」の拡張オーバーライドの操作 [の概要を取得します](../develop/extended-overrides.md)。
