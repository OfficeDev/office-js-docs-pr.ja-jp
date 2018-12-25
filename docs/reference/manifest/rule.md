---
title: マニフェスト ファイルの Rule 要素
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: ce7763ecb4ef81587ccacbd4090a6f412baf99b2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433116"
---
# <a name="rule-element"></a>Rule 要素

このコンテキスト メール アドインに対して評価する必要のあるアクティブ化ルールを指定します。

**アドインの種類:** メール コンテキスト アドイン

## <a name="contained-in"></a>次に含まれる

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

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
| **ItemClass** | いいえ | 照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](https://docs.microsoft.com/outlook/add-ins/activation-rules)」をご覧ください。 |
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
| **EntityType** | はい | このルールが true と評価されるために見つける必要のあるエンティティの型を指定します。`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress`、または `Contact` のいずれかになります。 |
| **RegExFilter** | いいえ | このエンティティに対してアクティブ化を実行するための正規表現を指定します。 |
| **FilterName** | いいえ | 正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。 |
| **IgnoreCase** | いいえ | **RegExFilter** 属性で指定した正規表現の実行時に、大文字と小文字の違いを無視するように指定します。 |
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
| **PropertyName** | はい | 正規表現の評価対象となるプロパティの名前を指定します。`Subject`、`BodyAsPlaintext`、`BodyAsHTML`、または `SenderSMTPAddress` のいずれかになります。 |
| **IgnoreCase** | いいえ | 正規表現の実行時に大文字と小文字の違いを無視するように指定します。 |
| **Highlight** | いいえ | **注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するテキストを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。 |

### <a name="example"></a>使用例

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。

### <a name="attributes"></a>属性

| 属性 | 必須 | 説明 |
|:-----|:-----|:-----|
| **Mode** | はい | このルール コレクションの評価時に使用する論理演算子を指定します。次のいずれかを指定できます。`And` または `Or` のどちらかになります。 |

### <a name="example"></a>使用例

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>関連項目

- [Outlook アドインのアクティブ化ルール](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)