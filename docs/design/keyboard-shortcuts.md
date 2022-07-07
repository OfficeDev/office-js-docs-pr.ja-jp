---
title: Office アドインのカスタム キーボード ショートカット
description: Office アドインにカスタム キーボード ショートカット (キーの組み合わせとも呼ばれる) を追加する方法について説明します。
ms.date: 11/22/2021
localization_priority: Normal
ms.openlocfilehash: 5e813e1f4af040bb546f60eb2db40862ba1a237e
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659984"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Office アドインにカスタム キーボード ショートカットを追加する

キーの組み合わせとも呼ばれるキーボード ショートカットを使用すると、アドインのユーザーがより効率的に作業できるようになります。 また、キーボード ショートカットを使用すると、マウスの代替手段を提供することで、障穣のあるユーザーに対するアドインのアクセシビリティも向上します。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> キーボード ショートカットが既に有効になっているアドインの作業バージョンから始めるには、サンプル [の Excel キーボード ショートカット](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts)を複製して実行します。 独自のアドインにキーボード ショートカットを追加する準備ができたら、この記事に進んでください。

アドインにキーボード ショートカットを追加するには、3 つの手順があります。

1. [アドインのマニフェストを構成します](#configure-the-manifest)。
1. [ショートカット JSON ファイルを作成または編集して、](#create-or-edit-the-shortcuts-json-file) アクションとそのキーボード ショートカットを定義します。
1. [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) API の [1 つ以上のランタイム呼び出しを追加](#create-a-mapping-of-actions-to-their-functions)して、関数を各アクションにマップします。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストには小さな変更が 2 つあります。 1 つは、アドインで共有ランタイムを使用できるようにすることです。もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルを指す方法です。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタム キーボード ショートカットを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、 [共有ランタイムを使用するようにアドインを構成します](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>マッピング ファイルをマニフェストにリンクする

マニフェスト内の要素の直 *下* (内部ではない) **\<VersionOverrides\>** に [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 要素を追加します。 後の `Url` 手順で作成するプロジェクト内の JSON ファイルの完全な URL に属性を設定します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>ショートカット JSON ファイルを作成または編集する

プロジェクトに JSON ファイルを作成します。 ファイルのパスが [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 要素の属性に`Url`指定した場所と一致していることを確認します。 このファイルでは、キーボード ショートカットと、それらが呼び出すアクションについて説明します。

1. JSON ファイル内には、2 つの配列があります。 actions 配列には、呼び出されるアクションを定義するオブジェクトが含まれます。ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。 次に例を示します。
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

    JSON オブジェクトの詳細については、「 [アクション オブジェクトを構築する](#construct-the-action-objects) 」と「 [ショートカット オブジェクトを構築する」を](#construct-the-shortcut-objects)参照してください。 ショートカット JSON の完全なスキーマは、 [extended-manifest.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

    > [!NOTE]
    > この記事では、"Ctrl" の代わりに "CONTROL" を使用できます。

    後の手順では、アクション自体が記述する関数にマップされます。 この例では、後で SHOWTASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップし、HIDETASKPANE をメソッドを呼び出す関数にマップします `Office.addin.hide` 。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>アクションの関数へのマッピングを作成する

1. プロジェクトで、要素内の HTML ページによって読み込まれた JavaScript ファイルを **\<FunctionFile\>** 開きます。
1. JavaScript ファイルで [Office.actions.associate API を](/javascript/api/office/office.actions#office-office-actions-associate-member) 使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。 次の JavaScript をファイルに追加します。 コードについては、次の点に注意してください。

    - 最初のパラメーターは、JSON ファイルからのアクションの 1 つです。
    - 2 番目のパラメーターは、ユーザーが JSON ファイル内のアクションにマップされているキーの組み合わせを押したときに実行される関数です。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. この例を続行するには、最初のパラメーターとして使用 `'SHOWTASKPANE'` します。
1. 関数の本体には、 [Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) メソッドを使用してアドインの作業ウィンドウを開きます。 完了すると、コードは次のようになります。

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

1. [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)) を呼び出す関数にアクションをマップ`HIDETASKPANE`する関数の 2 番目の呼び出`Office.actions.associate`しを追加します。 次に例を示します。

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

前の手順に従うと、アドインは **Ctrl + Alt + Up キーと Ctrl + Alt +** **Down** キーを押して作業ウィンドウの表示を切り替えることができます。 同じ動作は、GitHub の Office アドイン PnP リポジトリの [Excel キーボード ショートカット](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) サンプルに示されています。

## <a name="details-and-restrictions"></a>詳細と制限

### <a name="construct-the-action-objects"></a>アクション オブジェクトを構築する

shortcuts.json の配列内のオブジェクトを指定する場合は、 `actions` 次のガイドラインを使用します。

- プロパティ名 `id` と `name` 必須です。
- この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。
- プロパティは `name` 、アクションを説明するユーザー フレンドリな文字列である必要があります。 これは、文字 A - Z、a - z、0 - 9、句読点 "-"、"_"、"+" の組み合わせである必要があります。
- プロパティは省略可能です。 現時点では、型のみが `ExecuteFunction` サポートされています。

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

ショートカット JSON の完全なスキーマは、 [extended-manifest.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

### <a name="construct-the-shortcut-objects"></a>ショートカット オブジェクトを作成する

shortcuts.json の配列内のオブジェクトを指定する場合は、 `shortcuts` 次のガイドラインを使用します。

- プロパティ名 `action`、 `key`および `default` 必須です。
- プロパティの `action` 値は文字列であり、アクション オブジェクト内のプロパティのいずれか `id` と一致する必要があります。
- このプロパティには `default` 、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、"+" の任意の組み合わせを指定できます。 (慣例により、これらのプロパティでは小文字は使用されません)。
- プロパティには `default` 、少なくとも 1 つの修飾子キー (Alt、Ctrl、Shift) と他の 1 つのキーの名前を含む必要があります。
- Shift を唯一の修飾子キーとして使用することはできません。 Shift と Alt または Ctrl を組み合わせます。
- Mac の場合は、Command 修飾子キーもサポートされています。
- Mac の場合、Alt は Option キーにマップされます。 Windows の場合、コマンドは Ctrl キーにマップされます。
- 標準キーボードで 2 つの文字が同じ物理キーにリンクされている場合、それらはプロパティ内の `default` シノニムです。たとえば、Alt + a と Alt + A は同じショートカットであるため、Ctrl + と Ctrl +\_ は "-" と "_" が同じ物理キーであるためです。
- "+" 文字は、キーの両側のキーが同時に押されることを示します。

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

ショートカット JSON の完全なスキーマは、 [extended-manifest.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!NOTE]
> キーヒントは、キーヒント (塗りつぶしの色 **Alt + H、H** を選択する Excel ショートカットなど) とも呼ばれ、Office アドインではサポートされていません。

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>他のアドインで使用されているキーの組み合わせを避ける

Office で既に使用されているキーボード ショートカットは多数あります。 既に使用されているアドインのキーボード ショートカットを登録しないでください。ただし、既存のキーボード ショートカットをオーバーライドしたり、同じキーボード ショートカットを登録した複数のアドイン間の競合を処理したりする必要がある場合があります。

競合が発生した場合、ユーザーは、競合するキーボード ショートカットを初めて使用しようとしたときにダイアログ ボックスを表示します。 このダイアログに表示されるアドイン オプションのテキストは、ファイル内の `name` アクション オブジェクト `shortcuts.json` のプロパティに由来することに注意してください。

![1 つのショートカットに対して 2 つの異なるアクションを持つ競合モーダルを示す図。](../images/add-in-shortcut-conflict-modal.png)

ユーザーは、キーボード ショートカットで実行するアクションを選択できます。 選択を行った後、同じショートカットを今後使用するために設定が保存されます。 ショートカット設定は、プラットフォームごとにユーザーごとに保存されます。 ユーザーが自分の設定を変更する場合は **、検索ボックス** から **Office アドインのショートカット設定のリセット** コマンドを呼び出すことができます。 このコマンドを呼び出すと、ユーザーのアドインのショートカット設定がすべてクリアされ、競合するショートカットを次回使用しようとすると、競合ダイアログ ボックスが再度表示されます。

![Office アドインのショートカット設定のリセット アクションを示す Excel の [検索を指示する] ボックス。](../images/add-in-reset-shortcuts-action.png)

最適なユーザー エクスペリエンスを得るには、次の優れたプラクティスを使用して Excel との競合を最小限に抑えることが推奨されます。

- 次のパターンのキーボード ショートカットのみを使用します。**Ctrl + Shift + Alt +* x***。 *x* は他のキーです。
- その他のキーボード ショートカットが必要な場合は、 [Excel キーボード ショートカットの一覧](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)を確認し、アドインでキーボード ショートカットを使用しないでください。
- キーボード フォーカスがアドイン UI 内にある場合、 **Ctrl + Space キー** と **Ctrl + Shift + F10** は機能しません。これらは基本的なアクセシビリティ ショートカットであるためです。
- Windows または Mac コンピューターで、検索メニューで [Office アドインのショートカット設定のリセット] コマンドを使用できない場合、ユーザーはコンテキスト メニューからリボンをカスタマイズすることで、コマンドをリボンに手動で追加できます。

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>プラットフォームごとにキーボード ショートカットをカスタマイズする

ショートカットをプラットフォーム固有にカスタマイズできます。 次の各プラットフォームのショートカットを`shortcuts`カスタマイズするオブジェクトの例を次に示します。 `windows``mac``web` ショートカットごとにショートカット キーが `default` 必要であることに注意してください。

次の例では、 `default` キーは、指定されていないプラットフォームのフォールバック キーです。 指定されていないプラットフォームは Windows のみであるため `default` 、キーは Windows にのみ適用されます。

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

アドインで複数のロケールがサポートされている場合は、アクション オブジェクトのプロパティを `name` ローカライズする必要があります。 また、アドインでサポートされているロケールのアルファベットや書き込みシステムが異なり、そのためキーボードが異なる場合は、ショートカットもローカライズする必要があります。 キーボード ショートカット JSON をローカライズする方法については、「 [拡張オーバーライドをローカライズする](../develop/localization.md#localize-extended-overrides)」を参照してください。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>オーバーライドできないブラウザーのショートカット

Web でカスタム キーボード ショートカットを使用する場合、ブラウザーで使用される一部のキーボード ショートカットはアドインでオーバーライドできません。この一覧は進行中の作業です。 オーバーライドできない他の組み合わせが見つからない場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users"></a>特定のユーザーに対してカスタム キーボード ショートカットを有効にする

アドインを使用すると、ユーザーはアドインのアクションを別のキーボードの組み合わせに再割り当てできます。

> [!NOTE]
> このセクションで説明する API には [、KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) 要件セットが必要です。

[Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) メソッドを使用して、アドイン アクションにユーザーのカスタム キーボードの組み合わせを割り当てます。 メソッドは型 `{[actionId:string]: string|null}`のパラメーターを受け取り、アドイン `actionId`の拡張マニフェスト JSON で定義する必要があるアクション ID のサブセットです。 値は、ユーザーが推奨するキーの組み合わせです。 この値を指定することもできます `null`。これにより、その `actionId` カスタマイズが削除され、アドインの拡張マニフェスト JSON で定義されている既定のキーボードの組み合わせに戻ります。

ユーザーが Office にログインしている場合、カスタムの組み合わせはプラットフォームごとのユーザーのローミング設定に保存されます。 現在、匿名ユーザーのショートカットのカスタマイズはサポートされていません。

```javascript
const userCustomShortcuts = {
    SHOWTASKPANE:"CTRL+SHIFT+1", 
    HIDETASKPANE:"CTRL+SHIFT+2"
};
Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(function () {
        console.log("Successfully registered.");
    })
    .catch(function (ex) {
        if (ex.code == "InvalidOperation") {
            console.log("ActionId does not exist or shortcut combination is invalid.");
        }
    });
```

ユーザーに対して既に使用されているショートカットを確認するには、 [Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member) メソッドを呼び出します。 このメソッドは、ユーザーが指定したアクションを呼び出すために使用する必要がある現在のキーボードの組み合わせを表す型 `[actionId:string]:string|null}`のオブジェクトを返します。 値は、次の 3 つの異なるソースから取得できます。

- ショートカットとの競合があり、ユーザーがそのキーボードの組み合わせに対して別のアクション (ネイティブまたは別のアドイン) を使用することを選択した場合、返される `null` 値は、ショートカットがオーバーライドされ、ユーザーがそのアドイン アクションを呼び出すために現在使用できるキーボードの組み合わせがないためです。
- [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) メソッドを使用してショートカットがカスタマイズされている場合、返される値はカスタマイズされたキーボードの組み合わせになります。
- ショートカットがオーバーライドまたはカスタマイズされていない場合は、アドインの拡張マニフェスト JSON から値が返されます。

次に例を示します。

```javascript
Office.actions.getShortcuts()
    .then(function (userShortcuts) {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });

```

[「他のアドインで使用するキーの組み合わせを回避](#avoid-key-combinations-in-use-by-other-add-ins)する」で説明されているように、ショートカットでの競合を回避することをお勧めします。 1 つ以上のキーの組み合わせが既に使用されているかどうかを検出するには、それらを文字列の配列として [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member) メソッドに渡します。 このメソッドは、型のオブジェクトの配列の形式で既に使用されているキーの組み合わせを含むレポートを返します `{shortcut: string, inUse: boolean}`。 プロパティは `shortcut` 、"Ctrl + Shift + 1" などのキーの組み合わせです。 組み合わせが既に別のアクションに登録されている場合、 `inUse` プロパティは `true`. たとえば、「 `[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]` 」のように入力します。 次のコード スニペットは例です。

```javascript
const shortcuts = ["CTRL+SHIFT+1", "CTRL+SHIFT+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then(function (inUseArray) {
        const availableShortcuts = inUseArray.filter(function (shortcut) { return !shortcut.inUse; });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter(function (shortcut) { return shortcut.inUse; });
        console.log(usedShortcuts);
    });

```

## <a name="next-steps"></a>次の手順

- [Excel キーボード ショートカットの](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts)サンプル アドインを参照してください。
- マニフェストの拡張オーバーライドの操作に関するページで [、拡張オーバーライドの操作の概要を確認します](../develop/extended-overrides.md)。
