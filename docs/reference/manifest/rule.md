---
title: マニフェスト ファイルの Rule 要素
description: Rule 要素は、このコンテキストメールアドインに対して評価する必要があるアクティブ化ルールを指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 79b97f2e442e9d8ce59d17467161b5b9b7a7252d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641432"
---
# <a name="rule-element"></a>Rule 要素

このコンテキストメールアドインに対して評価する必要があるアクティブ化ルールを指定します。

**アドインの種類:** メール (コンテキスト)

## <a name="contained-in"></a>含まれる場所

- [OfficeApp](officeapp.md)
- [Extensionpoint](extensionpoint.md) ([**custompane** (非推奨)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **xsi:type** | はい | 定義されているルールの種類。 |

ルールの種類は、次のいずれかになります。

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>ItemIs ルール

選択したアイテムが指定した種類である場合に true と評価するルールを定義します。

### <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **ItemType** | はい | 照合するアイテムの種類を指定します。`Message` または `Appointment` になります。`Message` のアイテムの種類には、電子メール、会議出席依頼、会議出席依頼の返信、および会議のキャンセルが含まれます。 |
| **FormType** | いいえ ([ExtensionPoint](extensionpoint.md) 内)、いいえ ([OfficeApp](officeapp.md) 内) | アプリがアイテムの読み取りまたは編集フォームで表示されるかどうかを指定します。`Read`、`Edit` または `ReadOrEdit` のいずれかになります。`ExtensionPoint` 内の `Rule` で指定されている場合、この値は `Read` である必要があります。 |
| **ItemClass** | いいえ | 照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](../../outlook/activation-rules.md)」を参照してください。 |
| **IncludeSubClasses** | いいえ | アイテムが指定したメッセージ クラスのサブクラスである場合に、このルールは true と評価する必要があるかどうかを指定します。既定値は `false` です。 |

### <a name="example"></a>例

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment ルール

アイテムに添付ファイルがある場合に true と評価するルールを定義します。

### <a name="example"></a>例

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

指定したエンティティ型のテキストがアイテムの件名または本文に含まれている場合に true と評価するルールを定義します。

### <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **EntityType** | はい | このルールが true と評価するために見つける必要のあるエンティティの型を指定します。`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress`、または `Contact` のいずれかになります。 |
| **RegExFilter** | いいえ | このエンティティに対してアクティブ化を実行するための正規表現を指定します。 |
| **FilterName** | いいえ | 正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。 |
| **IgnoreCase** | いいえ | **RegExFilter** 属性で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。 |
| **Highlight** | いいえ | **注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するエンティティを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。 |

### <a name="example"></a>使用例

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

アイテムの指定したプロパティの中を検索し、指定した正規表現と一致するものがある場合に true と評価するルールを定義します。

### <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **RegExName** | はい | アドインのコードで参照できるように、正規表現の名前を指定します。 |
| **RegExValue** | はい | メール アドインを表示するかどうかを判断するために評価する正規表現を指定します。 |
| **PropertyName** | はい | 正規表現の評価対象となるプロパティの名前を指定します。`Subject`、`BodyAsPlaintext`、`BodyAsHTML`、または `SenderSMTPAddress` のいずれかになります。<br/><br/>`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。 HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。<br/><br/>`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。<br/><br/>**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。|
| **IgnoreCase** | いいえ | **RegExName** 属性で指定された正規表現の一致で大文字と小文字の違いを無視するかどうかを指定します。 |
| **Highlight** | いいえ | クライアントが一致するテキストを強調表示にする方法を指定します。 この属性は、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用できます。 `all` または `none` のいずれかになります。 指定のない場合、既定値は `all` に設定されます。<br/><br/>**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。
|

### <a name="example"></a>例

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。

### <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **Mode** | はい | このルール コレクションの評価時に使用する論理演算子を指定します。次のいずれかを指定できます。`And` または `Or` のどちらかになります。 |

### <a name="example"></a>例

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>関連項目

- [Outlook アドインのアクティブ化ルール](../../outlook/activation-rules.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
