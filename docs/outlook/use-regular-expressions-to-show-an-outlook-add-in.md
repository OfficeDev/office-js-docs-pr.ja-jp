---
title: 正規表現アクティブ化ルールを使用してアドインを表示する
description: Outlook コンテキスト アドインで正規表現アクティブ化ルールを使用する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: b697f1b0a4d20254986a7aa10a5cc7f25dbdd887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605242"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a><span data-ttu-id="c1cb7-103">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="c1cb7-103">Use regular expression activation rules to show an Outlook add-in</span></span>

<span data-ttu-id="c1cb7-104">メッセージの特定のフィールドで一致がある場合に[コンテキスト アドイン](contextual-outlook-add-ins.md)をアクティブ化するように正規表現ルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-104">You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message.</span></span> <span data-ttu-id="c1cb7-105">コンテキスト アドインは閲覧モードでのみアクティブになります。Outlook ではユーザーがアイテムを作成しているときにはコンテキスト アドインはアクティブになりません。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-105">Contextual add-ins activate only in read mode, Outlook does not activate contextual add-ins when the user is composing an item.</span></span> <span data-ttu-id="c1cb7-106">Outlook がアドインをアクティブにしない他のシナリオもあります。たとえば、アイテムが Information Rights Management (IRM) で保護されている場合です。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-106">There are also other scenarios where Outlook does not activate add-ins, for example, items protected by Information Rights Management (IRM).</span></span> <span data-ttu-id="c1cb7-107">詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-107">For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="c1cb7-108">アドイン XML マニフェストでは、[ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ルールまたは [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールの一部として正規表現を指定することができます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-108">You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule in the add-in XML manifest.</span></span> <span data-ttu-id="c1cb7-109">ルールは [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) 拡張点で指定されます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-109">The rules are specified in a [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) extension point.</span></span>

<span data-ttu-id="c1cb7-110">Outlook では、クライアント コンピューターのブラウザーで使用する JavaScript インタープリターのルールに基づいて正規表現を評価します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-110">Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer.</span></span> <span data-ttu-id="c1cb7-111">Outlook では、すべての XML プロセッサでもサポートされているものと同じ特殊文字リストをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-111">Outlook supports the same list of special characters that all XML processors also support.</span></span> <span data-ttu-id="c1cb7-112">次の表は、このような特殊文字を示しています。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-112">The following table lists these special characters.</span></span> <span data-ttu-id="c1cb7-113">これらの文字は、次の表に示すとおり、該当する文字にエスケープ シーケンスを指定すると正規表現で使用できます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-113">You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.</span></span>

<br/>

|<span data-ttu-id="c1cb7-114">文字</span><span class="sxs-lookup"><span data-stu-id="c1cb7-114">Character</span></span>|<span data-ttu-id="c1cb7-115">説明</span><span class="sxs-lookup"><span data-stu-id="c1cb7-115">Description</span></span>|<span data-ttu-id="c1cb7-116">使用するエスケープ シーケンス</span><span class="sxs-lookup"><span data-stu-id="c1cb7-116">Escape sequence to use</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="c1cb7-117">二重引用符</span><span class="sxs-lookup"><span data-stu-id="c1cb7-117">Double quotation mark</span></span>|`&quot;`|
|`&`|<span data-ttu-id="c1cb7-118">アンパサンド</span><span class="sxs-lookup"><span data-stu-id="c1cb7-118">Ampersand</span></span>|`&amp;`|
|`'`|<span data-ttu-id="c1cb7-119">アポストロフィ</span><span class="sxs-lookup"><span data-stu-id="c1cb7-119">Apostrophe</span></span>|`&apos;`|
|`<`|<span data-ttu-id="c1cb7-120">より小さい</span><span class="sxs-lookup"><span data-stu-id="c1cb7-120">Less-than sign</span></span>|`&lt;`|
|`>`|<span data-ttu-id="c1cb7-121">より大きい</span><span class="sxs-lookup"><span data-stu-id="c1cb7-121">Greater-than sign</span></span>|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="c1cb7-122">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="c1cb7-122">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="c1cb7-123">`ItemHasRegularExpressionMatch` ルールはサポートされているプロパティの特定の値に基づいてアドインのアクティブ化を制御するのに便利です。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-123">An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property.</span></span> <span data-ttu-id="c1cb7-124">`ItemHasRegularExpressionMatch` ルールには以下の属性があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-124">The `ItemHasRegularExpressionMatch` rule has the following attributes.</span></span>

<br/>

|<span data-ttu-id="c1cb7-125">属性名</span><span class="sxs-lookup"><span data-stu-id="c1cb7-125">Attribute name</span></span>|<span data-ttu-id="c1cb7-126">説明</span><span class="sxs-lookup"><span data-stu-id="c1cb7-126">Description</span></span>|
|:-----|:-----|
|`RegExName`|<span data-ttu-id="c1cb7-127">アドインのコードで参照できるように、正規表現の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-127">Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.</span></span>|
|`RegExValue`|<span data-ttu-id="c1cb7-128">アドインを表示するかどうかを判断するために評価する正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-128">Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.</span></span>|
|`PropertyName`|<span data-ttu-id="c1cb7-129">正規表現の評価対象となるプロパティの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-129">Specifies the name of the property that the regular expression will be evaluated against.</span></span> <span data-ttu-id="c1cb7-130">有効な値は `BodyAsHTML`、`BodyAsPlaintext`、`SenderSMTPAddress`、`Subject` です。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-130">The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.</span></span><br/><br/><span data-ttu-id="c1cb7-131">`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-131">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="c1cb7-132">HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-132">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="c1cb7-133">`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-133">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="c1cb7-134">**注:** `Rule` 要素に `Highlight` 属性を指定した場合は、`BodyAsPlaintext` に `PropertyName` 属性を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-134">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span>|
|`IgnoreCase`|<span data-ttu-id="c1cb7-135">`RegExName` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-135">Specifies whether to ignore case when matching the regular expression specified by `RegExName`.</span></span>|
| `Highlight` | <span data-ttu-id="c1cb7-136">クライアントが一致するテキストを強調表示にする方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-136">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="c1cb7-137">この要素は、`ExtensionPoint` 要素内の `Rule` 要素にのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-137">This element can only be applied to `Rule` elements within `ExtensionPoint` elements.</span></span> <span data-ttu-id="c1cb7-138">`all` または `none` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-138">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="c1cb7-139">指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-139">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="c1cb7-140">**注:** `Rule` 要素に `Highlight` 属性を指定した場合は、`BodyAsPlaintext` に `PropertyName` 属性を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-140">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span> |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a><span data-ttu-id="c1cb7-141">ルールで正規表現を使用する場合のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c1cb7-141">Best practices for using regular expressions in rules</span></span>

<span data-ttu-id="c1cb7-142">正規表現を使用する場合は、次の点に特に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-142">Pay special attention to the following when you use regular expressions:</span></span>

- <span data-ttu-id="c1cb7-143">アイテムの本文に `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-143">If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item.</span></span> <span data-ttu-id="c1cb7-144">`.*` などの正規表現を使用してアイテムの本文全体を取得しようとしても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-144">Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.</span></span>
- <span data-ttu-id="c1cb7-145">あるブラウザーで返されたプレーンテキストの本文は、別のブラウザーではわずかに異なることがあります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-145">The plain text body returned on one browser can be different in subtle ways on another.</span></span> <span data-ttu-id="c1cb7-146">`BodyAsPlaintext` を `PropertyName` 属性として `ItemHasRegularExpressionMatch` ルールを使用する場合は、アドインのサポート対象であるすべてのブラウザーで正規表現をテストします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-146">If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.</span></span>

    <span data-ttu-id="c1cb7-147">さまざまなブラウザーがさまざまな方法で選択したアイテムの本文を取得するため、使用している正規表現が、本文の一部として返される可能性がある微妙な違いをサポートしていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-147">Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text.</span></span> <span data-ttu-id="c1cb7-148">たとえば、アイテムの本文を取得するために、Internet Explorer 9 などのブラウザーでは DOM の `innerText` プロパティを使用し、Firefox などのその他のブラウザーでは `.textContent()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-148">For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item.</span></span> <span data-ttu-id="c1cb7-149">また、さまざまなブラウザーが異なる改行を返す場合があります。改行は、Internet Explorer では `\r\n`、Firefox および Chrome では `\n` です。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-149">Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome.</span></span> <span data-ttu-id="c1cb7-150">詳細については、「[W3C DOM の互換性 - HTML](https://quirksmode.org/dom/html/)」(W3C DOM の互換性 - HTML) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-150">For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).</span></span>

- <span data-ttu-id="c1cb7-151">アイテムの HTML 形式の本文は、Outlook リッチ クライアントと、Outlook on the web または Outlook モバイルとでは若干異なります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-151">The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook mobile.</span></span> <span data-ttu-id="c1cb7-152">正規表現を正確に定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-152">Define your regular expressions carefully.</span></span>

- <span data-ttu-id="c1cb7-p112">ホスト アプリケーション、デバイスの種類、または正規表現を適用するプロパティに応じて、ホストごとに、アクティブ化ルールとして正規表現を設計するときに認識しておく必要がある、ベスト プラクティスと制限事項が他にもあります。詳細については、「 [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-p112">Depending on the host application, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the hosts that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.</span></span>

### <a name="examples"></a><span data-ttu-id="c1cb7-155">例</span><span class="sxs-lookup"><span data-stu-id="c1cb7-155">Examples</span></span>

<span data-ttu-id="c1cb7-156">次の `ItemHasRegularExpressionMatch` ルールでは、大文字小文字に関係なく、送信者の SMTP メール アドレスが `@contoso` と一致した場合にアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-156">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

<span data-ttu-id="c1cb7-157">次の例では、`IgnoreCase` 属性を使用して同じ正規表現を指定しています。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-157">The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

<span data-ttu-id="c1cb7-158">次の `ItemHasRegularExpressionMatch` ルールでは、現在のアイテムの本文に株式銘柄コードが含まれている場合にアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-158">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="c1cb7-159">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="c1cb7-159">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="c1cb7-160">`ItemHasKnownEntity` ルールでは、選択したアイテムの件名または本文でのエンティティの存在に基づいてアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-160">An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item.</span></span> <span data-ttu-id="c1cb7-161">[EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) タイプはサポートされるエンティティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-161">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities.</span></span> <span data-ttu-id="c1cb7-162">`ItemHasKnownEntity` ルールに正規表現を適用すると、アクティブ化がエンティティの値のサブセット (特定の URL セットまたは、特定の市外局番の電話番号など) に基づく点で、利便性が増します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-162">Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).</span></span>

> [!NOTE]
> <span data-ttu-id="c1cb7-163">マニフェストに指定されている既定のロケールに関係なく、Outlook が抽出できるのは英語のエンティティ文字列だけです。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-163">Outlook can only extract entity strings in English regardless of the default locale specified in the manifest.</span></span> <span data-ttu-id="c1cb7-164">メッセージだけが `MeetingSuggestion` エンティティ タイプをサポートし、予定ではサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-164">Only messages support the `MeetingSuggestion` entity type; appointments do not.</span></span> <span data-ttu-id="c1cb7-165">**送信済みアイテム** フォルダーのアイテムからはエンティティを抽出できません。また、`ItemHasKnownEntity` ルールを使用して**送信済みアイテム** フォルダーにあるアイテムのにアドインを有効にすることもできません。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-165">You cannot extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.</span></span>

<span data-ttu-id="c1cb7-166">`ItemHasKnownEntity` ルールでは、以下の表にある属性をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-166">The `ItemHasKnownEntity` rule supports the attributes in the following table.</span></span> <span data-ttu-id="c1cb7-167">`ItemHasKnownEntity` ルールで正規表現の指定が任意の場合、エンティティ フィルターとして正規表現を使用するには、`RegExFilter` 属性と `FilterName` 属性の両方を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-167">Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.</span></span>

<br/>

|<span data-ttu-id="c1cb7-168">属性名</span><span class="sxs-lookup"><span data-stu-id="c1cb7-168">Attribute name</span></span>|<span data-ttu-id="c1cb7-169">説明</span><span class="sxs-lookup"><span data-stu-id="c1cb7-169">Description</span></span>|
|:-----|:-----|
|`EntityType`|<span data-ttu-id="c1cb7-170">このルールが `true` と評価するために見つける必要のあるエンティティの型を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-170">Specifies the type of entity that must be found for the rule to evaluate to `true`.</span></span> <span data-ttu-id="c1cb7-171">複数のルールを使用して複数のエンティティの型を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-171">Use multiple rules to specify multiple types of entities.</span></span>|
|`RegExFilter`|<span data-ttu-id="c1cb7-172">`EntityType` で指定されているエンティティのインスタンスをさらにフィルター処理する正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-172">Specifies a regular expression that further filters instances of the entity specified by `EntityType`.</span></span>|
|`FilterName`|<span data-ttu-id="c1cb7-173">`RegExFilter` で指定されている正規表現の名前を指定し、それ以降にコードでその正規表現を参照できるようにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-173">Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.</span></span>|
|`IgnoreCase`|<span data-ttu-id="c1cb7-174">`RegExFilter` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-174">Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.</span></span>|

### <a name="examples"></a><span data-ttu-id="c1cb7-175">例</span><span class="sxs-lookup"><span data-stu-id="c1cb7-175">Examples</span></span>

<span data-ttu-id="c1cb7-176">次の `ItemHasKnownEntity` ルールでは、現在のアイテムの件名または本文に URL が存在し、URL に文字列 `youtube` (大文字小文字は区別しない) が含まれている場合、常にアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-176">The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a><span data-ttu-id="c1cb7-177">コードでの正規表現の結果の使用</span><span class="sxs-lookup"><span data-stu-id="c1cb7-177">Using regular expression results in code</span></span>

<span data-ttu-id="c1cb7-178">現在のアイテムで次のメソッドを使用して、正規表現に一致するものを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-178">You can obtain matches to a regular expression by using the following methods on the current item:</span></span>

- <span data-ttu-id="c1cb7-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールと `ItemHasKnownEntity` ルールで指定されているすべての正規表現について、現在のアイテムで一致するものを返します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.</span></span>

- <span data-ttu-id="c1cb7-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールで指定されている特定された正規表現について、現在のアイテムで一致するものを返します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.</span></span>

- <span data-ttu-id="c1cb7-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) は、アドインの `ItemHasKnownEntity` ルールで指定されている正規表現について、一致するものを含むエンティティのインスタンス全体を返します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.</span></span>

<span data-ttu-id="c1cb7-182">正規表現が評価されると、配列オブジェクトに入れてアドインに一致が返されます。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-182">When the regular expressions are evaluated, the matches are returned to your add-in in an array object.</span></span> <span data-ttu-id="c1cb7-183">`getRegExMatches` については、そのオブジェクトに正規表現の名前の識別子があります。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-183">For `getRegExMatches`, that object has the identifier of the name of the regular expression.</span></span>

> [!NOTE]
> <span data-ttu-id="c1cb7-184">Outlook は、配列内の特定の順序で一致を返すわけではありません。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-184">Outlook does not return matches in any particular order in the array.</span></span> <span data-ttu-id="c1cb7-185">また、一致がこの配列と同じ順序で返されるとも想定できません。同じメールボックス内の同じアイテムにあるこれらの各クライアントで同じアドインを実行する場合においても同様です。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-185">Also, you should not assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.</span></span>

### <a name="examples"></a><span data-ttu-id="c1cb7-186">例</span><span class="sxs-lookup"><span data-stu-id="c1cb7-186">Examples</span></span>

<span data-ttu-id="c1cb7-187">`videoURL` という名前の正規表現を使用する `ItemHasRegularExpressionMatch` ルールを含めたコレクションの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-187">The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

<span data-ttu-id="c1cb7-188">次の例では、現在のアイテムの `getRegExMatches` を使用して、変数 `videos` を前の `ItemHasRegularExpressionMatch` ルールの結果に設定します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-188">The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

<span data-ttu-id="c1cb7-p119">このオブジェクトには、複数の一致が配列要素として格納されます。次のコード例は、 `reg1` という名前の正規表現に一致するものを反復処理して、HTML として表示する文字列を作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-p119">Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.</span></span>

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

<span data-ttu-id="c1cb7-191">`MeetingSuggestion` エンティティと `CampSuggestion` という正規表現を指定する `ItemHasKnownEntity` ルールの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-191">The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`.</span></span> <span data-ttu-id="c1cb7-192">現在選択されているアイテムに会議の提案が含まれ、件名または本文に `WonderCamp` という用語があると判明した場合、Outlook はこのアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-192">Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

<span data-ttu-id="c1cb7-193">次のコード例では、現在のアイテムの `getFilteredEntitiesByName` を使用して変数 `suggestions` を設定し、前の `ItemHasKnownEntity` ルールで検出された会議の提案の配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-193">The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a><span data-ttu-id="c1cb7-194">関連項目</span><span class="sxs-lookup"><span data-stu-id="c1cb7-194">See also</span></span>

- <span data-ttu-id="c1cb7-195">[Outlook アドイン: Contoso 社の注文番号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - 正規表現の一致に基づいてアクティブ化されるコンテキスト アドインのサンプル。</span><span class="sxs-lookup"><span data-stu-id="c1cb7-195">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.</span></span>
- [<span data-ttu-id="c1cb7-196">閲覧フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="c1cb7-196">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="c1cb7-197">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="c1cb7-197">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="c1cb7-198">Outlook アドインのアクティブ化と JavaScript API の制限</span><span class="sxs-lookup"><span data-stu-id="c1cb7-198">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="c1cb7-199">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="c1cb7-199">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- <span data-ttu-id="c1cb7-200">
  [.NET Framework での正規表現に関するベスト プラクティス](/dotnet/standard/base-types/best-practices)</span><span class="sxs-lookup"><span data-stu-id="c1cb7-200">[Best Practices for Regular Expressions in the .NET Framework](/dotnet/standard/base-types/best-practices)</span></span>
