---
title: 正規表現アクティブ化ルールを使用してアドインを表示する
description: Outlook コンテキスト アドインで正規表現アクティブ化ルールを使用する方法について説明します。
ms.date: 07/28/2020
ms.localizationpriority: medium
ms.openlocfilehash: f56d973ed3470b70bdfe834f9adc8a15a7623f0b
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149591"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>正規表現アクティブ化ルールを使用して Outlook アドインを表示する

メッセージの特定のフィールドで一致がある場合に[コンテキスト アドイン](contextual-outlook-add-ins.md)をアクティブ化するように正規表現ルールを指定します。 コンテキスト アドインは閲覧モードでのみアクティブになります。Outlook ではユーザーがアイテムを作成しているときにはコンテキスト アドインはアクティブになりません。 また、デジタル署名されたOutlookなど、アドインをアクティブ化しないシナリオも存在します。 詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。

アドイン XML マニフェストでは、[ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ルールまたは [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールの一部として正規表現を指定することができます。 ルールは [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) 拡張点で指定されます。

Outlook では、クライアント コンピューターのブラウザーで使用する JavaScript インタープリターのルールに基づいて正規表現を評価します。 Outlook では、すべての XML プロセッサでもサポートされているものと同じ特殊文字リストをサポートしています。 次の表は、このような特殊文字を示しています。 これらの文字は、次の表に示すとおり、該当する文字にエスケープ シーケンスを指定すると正規表現で使用できます。

<br/>

|文字|説明|使用するエスケープ シーケンス|
|:-----|:-----|:-----|
|`"`|二重引用符|`&quot;`|
|`&`|アンパサンド|`&amp;`|
|`'`|アポストロフィ|`&apos;`|
|`<`|より小さい|`&lt;`|
|`>`|より大きい|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

`ItemHasRegularExpressionMatch` ルールはサポートされているプロパティの特定の値に基づいてアドインのアクティブ化を制御するのに便利です。 `ItemHasRegularExpressionMatch` ルールには以下の属性があります。

<br/>

|属性名|説明|
|:-----|:-----|
|`RegExName`|アドインのコードで参照できるように、正規表現の名前を指定します。|
|`RegExValue`|アドインを表示するかどうかを判断するために評価する正規表現を指定します。|
|`PropertyName`|正規表現の評価対象となるプロパティの名前を指定します。 有効な値は `BodyAsHTML`、`BodyAsPlaintext`、`SenderSMTPAddress`、`Subject` です。<br/><br/>`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。 HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。<br/><br/>`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。<br/><br/>**注:** `Rule` 要素に `Highlight` 属性を指定した場合は、`BodyAsPlaintext` に `PropertyName` 属性を設定する必要があります。|
|`IgnoreCase`|`RegExName` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。|
| `Highlight` | クライアントが一致するテキストを強調表示にする方法を指定します。 この要素は、`ExtensionPoint` 要素内の `Rule` 要素にのみ適用できます。 `all` または `none` のいずれかになります。 指定のない場合、既定値は `all` に設定されます。<br/><br/>**注:** `Rule` 要素に `Highlight` 属性を指定した場合は、`BodyAsPlaintext` に `PropertyName` 属性を設定する必要があります。 |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>ルールで正規表現を使用する場合のベスト プラクティス

正規表現を使用する場合は、次の点に特に注意してください。

- アイテムの本文に `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。 `.*` などの正規表現を使用してアイテムの本文全体を取得しようとしても、期待する結果が返されないことがあります。
- あるブラウザーで返されたプレーンテキストの本文は、別のブラウザーではわずかに異なることがあります。 `BodyAsPlaintext` を `PropertyName` 属性として `ItemHasRegularExpressionMatch` ルールを使用する場合は、アドインのサポート対象であるすべてのブラウザーで正規表現をテストします。

    さまざまなブラウザーがさまざまな方法で選択したアイテムの本文を取得するため、使用している正規表現が、本文の一部として返される可能性がある微妙な違いをサポートしていることを確認する必要があります。 たとえば、アイテムの本文を取得するために、Internet Explorer 9 などのブラウザーでは DOM の `innerText` プロパティを使用し、Firefox などのその他のブラウザーでは `.textContent()` メソッドを使用します。 また、さまざまなブラウザーが異なる改行を返す場合があります。改行は、Internet Explorer では `\r\n`、Firefox および Chrome では `\n` です。 詳細については、「[W3C DOM の互換性 - HTML](https://quirksmode.org/dom/html/)」(W3C DOM の互換性 - HTML) を参照してください。

- アイテムの HTML 形式の本文は、Outlook リッチ クライアントと、Outlook on the web または Outlook モバイルとでは若干異なります。 正規表現を正確に定義する必要があります。

- 正規表現が適用されている Outlook クライアント、デバイスの種類、またはプロパティに応じて、アクティブ化ルールとして正規表現を設計するときに注意する必要がある各クライアントの他のベスト プラクティスと制限があります。 詳細については、「 [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」を参照してください。

### <a name="examples"></a>例

次の `ItemHasRegularExpressionMatch` ルールでは、大文字小文字に関係なく、送信者の SMTP メール アドレスが `@contoso` と一致した場合にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

次の例では、`IgnoreCase` 属性を使用して同じ正規表現を指定しています。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

次の `ItemHasRegularExpressionMatch` ルールでは、現在のアイテムの本文に株式銘柄コードが含まれている場合にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

`ItemHasKnownEntity` ルールでは、選択したアイテムの件名または本文でのエンティティの存在に基づいてアドインをアクティブにします。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) タイプはサポートされるエンティティを定義します。 `ItemHasKnownEntity` ルールに正規表現を適用すると、アクティブ化がエンティティの値のサブセット (特定の URL セットまたは、特定の市外局番の電話番号など) に基づく点で、利便性が増します。

> [!NOTE]
> マニフェストに指定されている既定のロケールに関係なく、Outlook が抽出できるのは英語のエンティティ文字列だけです。 メッセージだけが `MeetingSuggestion` エンティティ タイプをサポートし、予定ではサポートしていません。 **送信済みアイテム** フォルダーのアイテムからはエンティティを抽出できません。また、`ItemHasKnownEntity` ルールを使用して **送信済みアイテム** フォルダーにあるアイテムのにアドインを有効にすることもできません。

`ItemHasKnownEntity` ルールでは、以下の表にある属性をサポートしています。 `ItemHasKnownEntity` ルールで正規表現の指定が任意の場合、エンティティ フィルターとして正規表現を使用するには、`RegExFilter` 属性と `FilterName` 属性の両方を指定する必要があります。

<br/>

|属性名|説明|
|:-----|:-----|
|`EntityType`|このルールが `true` と評価するために見つける必要のあるエンティティの型を指定します。 複数のルールを使用して複数のエンティティの型を指定します。|
|`RegExFilter`|`EntityType` で指定されているエンティティのインスタンスをさらにフィルター処理する正規表現を指定します。|
|`FilterName`|`RegExFilter` で指定されている正規表現の名前を指定し、それ以降にコードでその正規表現を参照できるようにします。|
|`IgnoreCase`|`RegExFilter` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。|

### <a name="examples"></a>例

次の `ItemHasKnownEntity` ルールでは、現在のアイテムの件名または本文に URL が存在し、URL に文字列 `youtube` (大文字小文字は区別しない) が含まれている場合、常にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>コードでの正規表現の結果の使用

現在のアイテムに対して次のメソッドを使用して、正規表現に一致する文字列を取得できます。

- [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールと `ItemHasKnownEntity` ルールで指定されているすべての正規表現について、現在のアイテムで一致するものを返します。

- [getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールで指定されている特定された正規表現について、現在のアイテムで一致するものを返します。

- [getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasKnownEntity` ルールで指定されている正規表現について、一致するものを含むエンティティのインスタンス全体を返します。

正規表現が評価されると、配列オブジェクトに入れてアドインに一致が返されます。 `getRegExMatches` については、そのオブジェクトに正規表現の名前の識別子があります。

> [!NOTE]
> Outlook は、配列内の特定の順序で一致を返すわけではありません。 また、一致がこの配列と同じ順序で返されるとも想定できません。同じメールボックス内の同じアイテムにあるこれらの各クライアントで同じアドインを実行する場合においても同様です。

### <a name="examples"></a>例

`videoURL` という名前の正規表現を使用する `ItemHasRegularExpressionMatch` ルールを含めたコレクションの例を次に示します。

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

次の例では、現在のアイテムの `getRegExMatches` を使用して、変数 `videos` を前の `ItemHasRegularExpressionMatch` ルールの結果に設定します。

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

このオブジェクトには、複数の一致が配列要素として格納されます。次のコード例は、 `reg1` という名前の正規表現に一致するものを反復処理して、HTML として表示する文字列を作成する方法を示しています。

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

`MeetingSuggestion` エンティティと `CampSuggestion` という正規表現を指定する `ItemHasKnownEntity` ルールの例を次に示します。 現在選択されているアイテムに会議の提案が含まれ、件名または本文に `WonderCamp` という用語があると判明した場合、Outlook はこのアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

次のコード例では、現在のアイテムの `getFilteredEntitiesByName` を使用して変数 `suggestions` を設定し、前の `ItemHasKnownEntity` ルールで検出された会議の提案の配列を取得します。

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>関連項目

- [Outlook アドイン: Contoso 社の注文番号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - 正規表現の一致に基づいてアクティブ化されるコンテキスト アドインのサンプル。
- [閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)
- [Outlook アドインのアクティブ化ルール](activation-rules.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
- 
  [.NET Framework での正規表現に関するベスト プラクティス](/dotnet/standard/base-types/best-practices)
