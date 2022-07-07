---
title: コンテキスト Outlook アドイン
description: メッセージ自体から移動しなくてもそのメッセージに関連したタスクを開始できます。それにより、操作が簡単になると同時にユーザー エクスペリエンスが豊かになります。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 1a9e9e29862449911550c8c3a6dd44819816bd69
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660138"
---
# <a name="contextual-outlook-add-ins"></a>コンテキスト Outlook アドイン

コンテキスト アドインは、メッセージ内のテキストまたは予定に基づいてアクティブになる Outlook アドインです。コンテキスト アドインを使用すると、ユーザーはメッセージ自体から移動しなくてもそのメッセージに関連したタスクを開始できます。それにより、操作が簡単になると同時にユーザー エクスペリエンスが豊かになります。

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
> 注: `EmailAddress` と `Url` のエンティティ型は強調表示をサポートしていないため、コンテキスト アドインの起動に使用することはできません。ただし、追加のアクティブ化条件として、`RuleCollection` ルール型と組み合わせることはできます。

## <a name="how-to-launch-a-contextual-add-in"></a>コンテキスト アドインの起動方法

ユーザーは、既知のエンティティまたは開発者の正規表現のどちらかで、テキストを通じてコンテキスト アドインを起動します。通常、ユーザーはエンティティが強調表示されていることでコンテキスト アドインを特定します。次の例は、メッセージ内の強調表示の様子を示しています。このエンティティ (住所) は、青色と下線 (青色の点線) で示されています。ユーザーは、強調表示されたエンティティをクリックすることでコンテキスト アドインを起動します。 

**強調表示されているエンティティ (住所) が含まれるテキストの例**

![電子メール内で強調表示されているエンティティを表示します。](../images/outlook-detected-entity-highlight.png)
    
1 つのメッセージ内に複数のエンティティまたはコンテキスト アドインが存在する場合、ユーザー操作の規則がいくつかあります。

- エンティティが複数ある場合、ユーザーは対象のアドインを起動するために異なるエンティティをクリックする必要があります。
- エンティティが複数のアドインをアクティブにする場合、各アドインは新しいタブを開きます。ユーザーはタブを切り替えて、アドイン間の変更をします。たとえば、名前とアドレスは、電話のアドインとマップをトリガーするかもしれません。
- 1 つの文字列に複数のアドインをアクティブにする複数のエンティティが含まれる場合、文字列全体が強調表示され、その文字列をクリックすると、その文字列に関連付けられているすべてのアドインが別々のタブに表示されます。たとえば、レストランで会議を行う提案を説明する文字列によって、会議提案アドインとレストラン評価アドインをアクティブにできます。

## <a name="how-a-contextual-add-in-displays"></a>コンテキスト アドインの表示方法

アクティブ化されたコンテキスト アドインは、カード (エンティティの近くに現れる別ウィンドウ) で表示されます。通常、このカードはエンティティの下側に、できるだけ中央揃えになるように表示されます。エンティティの下側に十分な空間がない場合、カードはエンティティの上側に配置されます。次のスクリーンショットは、強調表示されたエンティティと、その下側のカード内でアクティブ化されたアドイン (Bing 地図) を示しています。

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
