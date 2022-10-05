---
title: コンテキスト Outlook アドイン
description: メッセージ自体から移動しなくてもそのメッセージに関連したタスクを開始できます。それにより、操作が簡単になると同時にユーザー エクスペリエンスが豊かになります。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 73a13787dac7a6e74db6b919cc01a6dd33d29ab5
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467026"
---
# <a name="contextual-outlook-add-ins"></a>コンテキスト Outlook アドイン

Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

コンテキスト アドインの例を次に示します。

- 住所を選択すると、その場所の地図が開きます。
- 文字列をクリックすると、会議提案アドインが開きます。
- 電話番号を選択すると、連絡先に追加されます。


> [!NOTE]
> 現在、Android および iOS 用の Outlook では、コンテキスト アドインをご利用いただけません。 今後、この機能が使用可能になる予定です。
>
> この機能のサポートは、要件セット 1.6 に導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="how-to-make-a-contextual-add-in"></a>コンテキスト アドインの作成方法

コンテキスト アドインのマニフェストには、`xsi:type` 属性が `DetectedEntity` に設定されている [ExtensionPoint](/javascript/api/manifest/extensionpoint#detectedentity) 要素が含まれている必要があります。 **\<ExtensionPoint\>** 要素内で、アドインは、それをアクティブ化できるエンティティまたは正規表現を指定します。 エンティティを指定する場合、そのエンティティは [Entities](/javascript/api/outlook/office.entities) オブジェクトのどのプロパティであってもかまいません。

そのため、アドイン マニフェストには、ルールの種類 **ItemHasKnownEntity** または **ItemHasRegularExpressionMatch** が含まれている必要があります。 次の例では、電話番号であるエンティティが検出されたメッセージに対してアドインをアクティブ化するように指定する方法を示します。

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

コンテキスト アドインをアカウントに関連付けると、強調表示された状態のエンティティまたは正規表現をユーザーがクリックするとコンテキスト アプリが自動的に起動します。 Outlook アドインでの正規表現について詳しくは、「[正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)」を参照してください。

コンテキスト アドインには、次のいくつかの制限があります。

- コンテキスト アドインを含めることができるのは読み取りアドインのみです (作成アドインは不可)。
- 強調表示されたエンティティの色は指定できません。
- 強調表示されていないエンティティは、コンテキスト アドインをカード内で起動することはありません。

強調表示されていないエンティティまたは正規表現はコンテキスト アドインを起動しないため、アドイン マニフェストには `Highlight` 属性が `all` に設定された `Rule` 要素を少なくとも 1 つは含んでいる必要があります。

> [!NOTE]
> The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.

## <a name="how-to-launch-a-contextual-add-in"></a>コンテキスト アドインの起動方法

A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity. 

**強調表示されているエンティティ (住所) が含まれるテキストの例**

![電子メール内で強調表示されているエンティティを表示します。](../images/outlook-detected-entity-highlight.png)
    
1 つのメッセージ内に複数のエンティティまたはコンテキスト アドインが存在する場合、ユーザー操作の規則がいくつかあります。

- エンティティが複数ある場合、ユーザーは対象のアドインを起動するために異なるエンティティをクリックする必要があります。
- エンティティが複数のアドインをアクティブにする場合、各アドインは新しいタブを開きます。ユーザーはタブを切り替えて、アドイン間の変更をします。たとえば、名前とアドレスは、電話のアドインとマップをトリガーするかもしれません。
- If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.

## <a name="how-a-contextual-add-in-displays"></a>コンテキスト アドインの表示方法

An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.

**カードに表示されるアドインの例**

![カード内のコンテキスト アプリを示す。](../images/outlook-detected-entity-card.png)

カードを閉じてアドインを終了するには、カードの外側で任意の場所をクリックします。

## <a name="current-contextual-add-ins"></a>現在のコンテキスト アドイン

次のコンテキスト アドインは、Outlook アドインを使用するユーザーに対して既定でインストールされます。

- Bing 地図
- 会議の候補

## <a name="see-also"></a>関連項目

- [Outlook アドイン: Contoso 社の注文番号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (正規表現の一致に基づいてアクティブ化されるコンテキスト アドインのサンプル)
- [初めて Outlook アドインを記述する](../quickstarts/outlook-quickstart.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Entities オブジェクト](/javascript/api/outlook/office.entities)
