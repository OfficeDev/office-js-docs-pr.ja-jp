---
title: マニフェスト ファイルの Rule 要素
description: Rule 要素は、このコンテキストメールアドインに対して評価する必要があるアクティブ化ルールを指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: c1f2be3371333bfd87e0693d02a9a5984c18317b
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253622"
---
# <a name="rule-element"></a><span data-ttu-id="ae8bf-103">Rule 要素</span><span class="sxs-lookup"><span data-stu-id="ae8bf-103">Rule element</span></span>

<span data-ttu-id="ae8bf-104">このコンテキストメールアドインに対して評価する必要があるアクティブ化ルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="ae8bf-105">**アドインの種類:** メール (コンテキスト)</span><span class="sxs-lookup"><span data-stu-id="ae8bf-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="ae8bf-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="ae8bf-106">Contained in</span></span>

- [<span data-ttu-id="ae8bf-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ae8bf-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="ae8bf-108">[Extensionpoint](extensionpoint.md) ([**custompane** (非推奨)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span><span class="sxs-lookup"><span data-stu-id="ae8bf-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="ae8bf-109">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-109">Attributes</span></span>

| <span data-ttu-id="ae8bf-110">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-110">Attribute</span></span> | <span data-ttu-id="ae8bf-111">必須</span><span class="sxs-lookup"><span data-stu-id="ae8bf-111">Required</span></span> | <span data-ttu-id="ae8bf-112">説明</span><span class="sxs-lookup"><span data-stu-id="ae8bf-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae8bf-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-113">**xsi:type**</span></span> | <span data-ttu-id="ae8bf-114">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-114">Yes</span></span> | <span data-ttu-id="ae8bf-115">定義されているルールの種類。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-115">The type of rule being defined.</span></span> |

<span data-ttu-id="ae8bf-116">ルールの種類は、次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="ae8bf-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="ae8bf-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="ae8bf-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="ae8bf-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="ae8bf-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ae8bf-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="ae8bf-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="ae8bf-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="ae8bf-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ae8bf-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="ae8bf-122">ItemIs ルール</span><span class="sxs-lookup"><span data-stu-id="ae8bf-122">ItemIs rule</span></span>

<span data-ttu-id="ae8bf-123">選択したアイテムが指定した種類である場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae8bf-124">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-124">Attributes</span></span>

| <span data-ttu-id="ae8bf-125">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-125">Attribute</span></span> | <span data-ttu-id="ae8bf-126">必須</span><span class="sxs-lookup"><span data-stu-id="ae8bf-126">Required</span></span> | <span data-ttu-id="ae8bf-127">説明</span><span class="sxs-lookup"><span data-stu-id="ae8bf-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae8bf-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-128">**ItemType**</span></span> | <span data-ttu-id="ae8bf-129">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-129">Yes</span></span> | <span data-ttu-id="ae8bf-p101">照合するアイテムの種類を指定します。`Message` または `Appointment` になります。`Message` のアイテムの種類には、電子メール、会議出席依頼、会議出席依頼の返信、および会議のキャンセルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="ae8bf-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-133">**FormType**</span></span> | <span data-ttu-id="ae8bf-134">いいえ ([ExtensionPoint](extensionpoint.md) 内)、いいえ ([OfficeApp](officeapp.md) 内)</span><span class="sxs-lookup"><span data-stu-id="ae8bf-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="ae8bf-p102">アプリがアイテムの読み取りまたは編集フォームで表示されるかどうかを指定します。`Read`、`Edit` または `ReadOrEdit` のいずれかになります。`ExtensionPoint` 内の `Rule` で指定されている場合、この値は `Read` である必要があります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="ae8bf-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-138">**ItemClass**</span></span> | <span data-ttu-id="ae8bf-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-139">No</span></span> | <span data-ttu-id="ae8bf-p103">照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](../../outlook/activation-rules.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="ae8bf-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="ae8bf-143">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-143">No</span></span> | <span data-ttu-id="ae8bf-144">アイテムが指定したメッセージ クラスのサブクラスである場合に、このルールは true と評価する必要があるかどうかを指定します。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae8bf-145">例</span><span class="sxs-lookup"><span data-stu-id="ae8bf-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="ae8bf-146">ItemHasAttachment ルール</span><span class="sxs-lookup"><span data-stu-id="ae8bf-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="ae8bf-147">アイテムに添付ファイルがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="ae8bf-148">例</span><span class="sxs-lookup"><span data-stu-id="ae8bf-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="ae8bf-149">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="ae8bf-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="ae8bf-150">指定したエンティティ型のテキストがアイテムの件名または本文に含まれている場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae8bf-151">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-151">Attributes</span></span>

| <span data-ttu-id="ae8bf-152">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-152">Attribute</span></span> | <span data-ttu-id="ae8bf-153">必須</span><span class="sxs-lookup"><span data-stu-id="ae8bf-153">Required</span></span> | <span data-ttu-id="ae8bf-154">説明</span><span class="sxs-lookup"><span data-stu-id="ae8bf-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae8bf-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-155">**EntityType**</span></span> | <span data-ttu-id="ae8bf-156">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-156">Yes</span></span> | <span data-ttu-id="ae8bf-p104">このルールが true と評価するために見つける必要のあるエンティティの型を指定します。`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress`、または `Contact` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="ae8bf-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-159">**RegExFilter**</span></span> | <span data-ttu-id="ae8bf-160">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-160">No</span></span> | <span data-ttu-id="ae8bf-161">このエンティティに対してアクティブ化を実行するための正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="ae8bf-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-162">**FilterName**</span></span> | <span data-ttu-id="ae8bf-163">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-163">No</span></span> | <span data-ttu-id="ae8bf-164">正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="ae8bf-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-165">**IgnoreCase**</span></span> | <span data-ttu-id="ae8bf-166">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-166">No</span></span> | <span data-ttu-id="ae8bf-167">**RegExFilter** 属性で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="ae8bf-168">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-168">**Highlight**</span></span> | <span data-ttu-id="ae8bf-169">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-169">No</span></span> | <span data-ttu-id="ae8bf-p105">**注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するエンティティを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae8bf-174">使用例</span><span class="sxs-lookup"><span data-stu-id="ae8bf-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="ae8bf-175">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="ae8bf-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="ae8bf-176">アイテムの指定したプロパティの中を検索し、指定した正規表現と一致するものがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae8bf-177">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-177">Attributes</span></span>

| <span data-ttu-id="ae8bf-178">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-178">Attribute</span></span> | <span data-ttu-id="ae8bf-179">必須</span><span class="sxs-lookup"><span data-stu-id="ae8bf-179">Required</span></span> | <span data-ttu-id="ae8bf-180">説明</span><span class="sxs-lookup"><span data-stu-id="ae8bf-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae8bf-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-181">**RegExName**</span></span> | <span data-ttu-id="ae8bf-182">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-182">Yes</span></span> | <span data-ttu-id="ae8bf-183">アドインのコードで参照できるように、正規表現の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="ae8bf-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-184">**RegExValue**</span></span> | <span data-ttu-id="ae8bf-185">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-185">Yes</span></span> | <span data-ttu-id="ae8bf-186">メール アドインを表示するかどうかを判断するために評価する正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="ae8bf-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-187">**PropertyName**</span></span> | <span data-ttu-id="ae8bf-188">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-188">Yes</span></span> | <span data-ttu-id="ae8bf-p106">正規表現の評価対象となるプロパティの名前を指定します。`Subject`、`BodyAsPlaintext`、`BodyAsHTML`、または `SenderSMTPAddress` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="ae8bf-191">`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="ae8bf-192">HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="ae8bf-193">`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="ae8bf-194">**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="ae8bf-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-195">**IgnoreCase**</span></span> | <span data-ttu-id="ae8bf-196">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-196">No</span></span> | <span data-ttu-id="ae8bf-197">**RegExName** 属性で指定された正規表現の一致で大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="ae8bf-198">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-198">**Highlight**</span></span> | <span data-ttu-id="ae8bf-199">いいえ</span><span class="sxs-lookup"><span data-stu-id="ae8bf-199">No</span></span> | <span data-ttu-id="ae8bf-200">クライアントが一致するテキストを強調表示にする方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="ae8bf-201">この属性は、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="ae8bf-202">`all` または `none` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="ae8bf-203">指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="ae8bf-204">**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="ae8bf-205">例</span><span class="sxs-lookup"><span data-stu-id="ae8bf-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="ae8bf-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ae8bf-206">RuleCollection</span></span>

<span data-ttu-id="ae8bf-207">ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae8bf-208">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-208">Attributes</span></span>

| <span data-ttu-id="ae8bf-209">属性</span><span class="sxs-lookup"><span data-stu-id="ae8bf-209">Attribute</span></span> | <span data-ttu-id="ae8bf-210">必須</span><span class="sxs-lookup"><span data-stu-id="ae8bf-210">Required</span></span> | <span data-ttu-id="ae8bf-211">説明</span><span class="sxs-lookup"><span data-stu-id="ae8bf-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae8bf-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="ae8bf-212">**Mode**</span></span> | <span data-ttu-id="ae8bf-213">はい</span><span class="sxs-lookup"><span data-stu-id="ae8bf-213">Yes</span></span> | <span data-ttu-id="ae8bf-p109">このルール コレクションの評価時に使用する論理演算子を指定します。次のいずれかを指定できます。`And` または `Or` のどちらかになります。</span><span class="sxs-lookup"><span data-stu-id="ae8bf-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae8bf-216">例</span><span class="sxs-lookup"><span data-stu-id="ae8bf-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="ae8bf-217">関連項目</span><span class="sxs-lookup"><span data-stu-id="ae8bf-217">See also</span></span>

- [<span data-ttu-id="ae8bf-218">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="ae8bf-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="ae8bf-219">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="ae8bf-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="ae8bf-220">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="ae8bf-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
