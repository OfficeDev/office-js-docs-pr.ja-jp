---
title: マニフェスト ファイルの Rule 要素
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: 2c5ae07e5d0a3c9c8979abcada3d758c415e2e59
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457454"
---
# <a name="rule-element"></a><span data-ttu-id="212b6-102">Rule 要素</span><span class="sxs-lookup"><span data-stu-id="212b6-102">Rule element</span></span>

<span data-ttu-id="212b6-103">このコンテキスト メール アドインに対して評価する必要のあるアクティブ化ルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="212b6-104">**アドインの種類:** メール コンテキスト アドイン</span><span class="sxs-lookup"><span data-stu-id="212b6-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="212b6-105">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="212b6-105">Contained in</span></span>

- [<span data-ttu-id="212b6-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="212b6-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="212b6-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="212b6-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="212b6-108">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-108">Attributes</span></span>

| <span data-ttu-id="212b6-109">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-109">Attribute</span></span> | <span data-ttu-id="212b6-110">必須</span><span class="sxs-lookup"><span data-stu-id="212b6-110">Required</span></span> | <span data-ttu-id="212b6-111">説明</span><span class="sxs-lookup"><span data-stu-id="212b6-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="212b6-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="212b6-112">**xsi:type**</span></span> | <span data-ttu-id="212b6-113">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-113">Yes</span></span> | <span data-ttu-id="212b6-114">定義されているルールの種類。</span><span class="sxs-lookup"><span data-stu-id="212b6-114">The type of rule being defined.</span></span> |

<span data-ttu-id="212b6-115">ルールの種類は、次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="212b6-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="212b6-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="212b6-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="212b6-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="212b6-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="212b6-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="212b6-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="212b6-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="212b6-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="212b6-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="212b6-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="212b6-121">ItemIs ルール</span><span class="sxs-lookup"><span data-stu-id="212b6-121">ItemIs rule</span></span>

<span data-ttu-id="212b6-122">選択したアイテムが指定した種類である場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="212b6-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="212b6-123">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-123">Attributes</span></span>

| <span data-ttu-id="212b6-124">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-124">Attribute</span></span> | <span data-ttu-id="212b6-125">必須</span><span class="sxs-lookup"><span data-stu-id="212b6-125">Required</span></span> | <span data-ttu-id="212b6-126">説明</span><span class="sxs-lookup"><span data-stu-id="212b6-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="212b6-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="212b6-127">**ItemType**</span></span> | <span data-ttu-id="212b6-128">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-128">Yes</span></span> | <span data-ttu-id="212b6-p101">照合するアイテムの種類を指定します。`Message` または `Appointment` になります。`Message` のアイテムの種類には、電子メール、会議出席依頼、会議出席依頼の返信、および会議のキャンセルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="212b6-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="212b6-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="212b6-132">**FormType**</span></span> | <span data-ttu-id="212b6-133">いいえ ([ExtensionPoint](extensionpoint.md) 内)、いいえ ([OfficeApp](officeapp.md) 内)</span><span class="sxs-lookup"><span data-stu-id="212b6-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="212b6-p102">アプリがアイテムの読み取りまたは編集フォームで表示されるかどうかを指定します。`Read`、`Edit` または `ReadOrEdit` のいずれかになります。`ExtensionPoint` 内の `Rule` で指定されている場合、この値は `Read` である必要があります。</span><span class="sxs-lookup"><span data-stu-id="212b6-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="212b6-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="212b6-137">**ItemClass**</span></span> | <span data-ttu-id="212b6-138">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-138">No</span></span> | <span data-ttu-id="212b6-p103">照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](https://docs.microsoft.com/outlook/add-ins/activation-rules)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="212b6-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="212b6-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="212b6-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="212b6-142">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-142">No</span></span> | <span data-ttu-id="212b6-143">アイテムが指定したメッセージ クラスのサブクラスである場合に、このルールは true と評価する必要があるかどうかを指定します。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="212b6-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="212b6-144">例</span><span class="sxs-lookup"><span data-stu-id="212b6-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="212b6-145">ItemHasAttachment ルール</span><span class="sxs-lookup"><span data-stu-id="212b6-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="212b6-146">アイテムに添付ファイルがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="212b6-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="212b6-147">例</span><span class="sxs-lookup"><span data-stu-id="212b6-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="212b6-148">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="212b6-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="212b6-149">指定したエンティティ型のテキストがアイテムの件名または本文に含まれている場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="212b6-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="212b6-150">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-150">Attributes</span></span>

| <span data-ttu-id="212b6-151">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-151">Attribute</span></span> | <span data-ttu-id="212b6-152">必須</span><span class="sxs-lookup"><span data-stu-id="212b6-152">Required</span></span> | <span data-ttu-id="212b6-153">説明</span><span class="sxs-lookup"><span data-stu-id="212b6-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="212b6-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="212b6-154">**EntityType**</span></span> | <span data-ttu-id="212b6-155">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-155">Yes</span></span> | <span data-ttu-id="212b6-p104">このルールが true と評価されるために見つける必要のあるエンティティの型を指定します。`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress`、または `Contact` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="212b6-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="212b6-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="212b6-158">**RegExFilter**</span></span> | <span data-ttu-id="212b6-159">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-159">No</span></span> | <span data-ttu-id="212b6-160">このエンティティに対してアクティブ化を実行するための正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="212b6-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="212b6-161">**FilterName**</span></span> | <span data-ttu-id="212b6-162">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-162">No</span></span> | <span data-ttu-id="212b6-163">正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。</span><span class="sxs-lookup"><span data-stu-id="212b6-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="212b6-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="212b6-164">**IgnoreCase**</span></span> | <span data-ttu-id="212b6-165">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-165">No</span></span> | <span data-ttu-id="212b6-166">**RegExFilter** 属性で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-166">Specifies whether to ignore case when matching the regular expression specified by  **RegExFilter**.</span></span> |
| <span data-ttu-id="212b6-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="212b6-167">**Highlight**</span></span> | <span data-ttu-id="212b6-168">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-168">No</span></span> | <span data-ttu-id="212b6-p105">**注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するエンティティを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="212b6-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="212b6-173">使用例</span><span class="sxs-lookup"><span data-stu-id="212b6-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="212b6-174">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="212b6-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="212b6-175">アイテムの指定したプロパティの中を検索し、指定した正規表現と一致するものがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="212b6-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="212b6-176">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-176">Attributes</span></span>

| <span data-ttu-id="212b6-177">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-177">Attribute</span></span> | <span data-ttu-id="212b6-178">必須</span><span class="sxs-lookup"><span data-stu-id="212b6-178">Required</span></span> | <span data-ttu-id="212b6-179">説明</span><span class="sxs-lookup"><span data-stu-id="212b6-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="212b6-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="212b6-180">**RegExName**</span></span> | <span data-ttu-id="212b6-181">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-181">Yes</span></span> | <span data-ttu-id="212b6-182">アドインのコードで参照できるように、正規表現の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="212b6-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="212b6-183">**RegExValue**</span></span> | <span data-ttu-id="212b6-184">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-184">Yes</span></span> | <span data-ttu-id="212b6-185">メール アドインを表示するかどうかを判断するために評価する正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="212b6-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="212b6-186">**PropertyName**</span></span> | <span data-ttu-id="212b6-187">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-187">Yes</span></span> | <span data-ttu-id="212b6-p106">正規表現の評価対象となるプロパティの名前を指定します。`Subject`、`BodyAsPlaintext`、`BodyAsHTML`、または `SenderSMTPAddress` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="212b6-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="212b6-190">`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。</span><span class="sxs-lookup"><span data-stu-id="212b6-190">If you specify `BodyAsHTML`, Outlook applies the regular expression only if the item body is HTML, and otherwise Outlook returns no matches for that regular expression.</span></span> <span data-ttu-id="212b6-191">HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。</span><span class="sxs-lookup"><span data-stu-id="212b6-191">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="212b6-192">`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。</span><span class="sxs-lookup"><span data-stu-id="212b6-192">If you specify  `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="212b6-193">**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212b6-193">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="212b6-194">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="212b6-194">**IgnoreCase**</span></span> | <span data-ttu-id="212b6-195">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-195">No</span></span> | <span data-ttu-id="212b6-196">**RegExName** 属性で指定された正規表現の一致で大文字と小文字の違いを無視するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-196">Specifies whether to ignore case when matching the regular expression specified by  **RegExName**.</span></span> |
| <span data-ttu-id="212b6-197">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="212b6-197">**Highlight**</span></span> | <span data-ttu-id="212b6-198">いいえ</span><span class="sxs-lookup"><span data-stu-id="212b6-198">No</span></span> | <span data-ttu-id="212b6-199">クライアントが一致するテキストを強調表示にする方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="212b6-199">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="212b6-200">この属性は、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="212b6-200">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="212b6-201">`all` または `none` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="212b6-201">Can be one of the following constants:  `all` or `none`.</span></span> <span data-ttu-id="212b6-202">指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="212b6-202">If this value is not specified, the default value of the `all` property is used.</span></span><br/><br/><span data-ttu-id="212b6-203">**注:** **Rule** 要素に **Highlight** 属性を指定した場合は、**PropertyName** 属性を `BodyAsPlaintext` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212b6-203">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="212b6-204">例</span><span class="sxs-lookup"><span data-stu-id="212b6-204">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="212b6-205">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="212b6-205">RuleCollection</span></span>

<span data-ttu-id="212b6-206">ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。</span><span class="sxs-lookup"><span data-stu-id="212b6-206">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="212b6-207">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-207">Attributes</span></span>

| <span data-ttu-id="212b6-208">属性</span><span class="sxs-lookup"><span data-stu-id="212b6-208">Attribute</span></span> | <span data-ttu-id="212b6-209">必須</span><span class="sxs-lookup"><span data-stu-id="212b6-209">Required</span></span> | <span data-ttu-id="212b6-210">説明</span><span class="sxs-lookup"><span data-stu-id="212b6-210">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="212b6-211">**Mode**</span><span class="sxs-lookup"><span data-stu-id="212b6-211">**Mode**</span></span> | <span data-ttu-id="212b6-212">はい</span><span class="sxs-lookup"><span data-stu-id="212b6-212">Yes</span></span> | <span data-ttu-id="212b6-p109">このルール コレクションの評価時に使用する論理演算子を指定します。次のいずれかを指定できます。`And` または `Or` のどちらかになります。</span><span class="sxs-lookup"><span data-stu-id="212b6-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="212b6-215">使用例</span><span class="sxs-lookup"><span data-stu-id="212b6-215">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="212b6-216">関連項目</span><span class="sxs-lookup"><span data-stu-id="212b6-216">See also</span></span>

- [<span data-ttu-id="212b6-217">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="212b6-217">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="212b6-218">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="212b6-218">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="212b6-219">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="212b6-219">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)