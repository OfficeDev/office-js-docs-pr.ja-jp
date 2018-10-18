# <a name="rule-element"></a><span data-ttu-id="9bf81-101">Rule 要素</span><span class="sxs-lookup"><span data-stu-id="9bf81-101">Rule element</span></span>

<span data-ttu-id="9bf81-102">このコンテキスト メール アドインに対して評価する必要のあるアクティブ化ルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="9bf81-103">**アドインの種類:** メール コンテキスト アドイン</span><span class="sxs-lookup"><span data-stu-id="9bf81-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="9bf81-104">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="9bf81-104">Contained in:</span></span>

- [<span data-ttu-id="9bf81-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9bf81-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="9bf81-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="9bf81-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="9bf81-107">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-107">Attributes</span></span>

| <span data-ttu-id="9bf81-108">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-108">Attribute</span></span> | <span data-ttu-id="9bf81-109">必須</span><span class="sxs-lookup"><span data-stu-id="9bf81-109">Required</span></span> | <span data-ttu-id="9bf81-110">説明</span><span class="sxs-lookup"><span data-stu-id="9bf81-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="9bf81-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="9bf81-111">**xsi:type**</span></span> | <span data-ttu-id="9bf81-112">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-112">Yes</span></span> | <span data-ttu-id="9bf81-113">定義されているルールの種類。</span><span class="sxs-lookup"><span data-stu-id="9bf81-113">The type of rule being defined.</span></span> |

<span data-ttu-id="9bf81-114">ルールの種類は、次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="9bf81-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="9bf81-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="9bf81-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="9bf81-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="9bf81-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="9bf81-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="9bf81-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="9bf81-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="9bf81-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="9bf81-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="9bf81-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="9bf81-120">ItemIs ルール</span><span class="sxs-lookup"><span data-stu-id="9bf81-120">ItemIs rule</span></span>

<span data-ttu-id="9bf81-121">選択したアイテムが指定した種類である場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="9bf81-122">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-122">Attributes</span></span>

| <span data-ttu-id="9bf81-123">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-123">Attribute</span></span> | <span data-ttu-id="9bf81-124">必須</span><span class="sxs-lookup"><span data-stu-id="9bf81-124">Required</span></span> | <span data-ttu-id="9bf81-125">説明</span><span class="sxs-lookup"><span data-stu-id="9bf81-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="9bf81-126">**アイテムタイプ**</span><span class="sxs-lookup"><span data-stu-id="9bf81-126">**ItemType**</span></span> | <span data-ttu-id="9bf81-127">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-127">Yes</span></span> | <span data-ttu-id="9bf81-p101">照合するアイテムの種類を指定します。`Message` または `Appointment` になります。`Message` のアイテムの種類には、電子メール、会議出席依頼、会議出席依頼の返信、および会議のキャンセルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="9bf81-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="9bf81-131">**FormType**</span></span> | <span data-ttu-id="9bf81-132">いいえ ([ExtensionPoint](extensionpoint.md) 内)、いいえ ([OfficeApp](officeapp.md) 内)</span><span class="sxs-lookup"><span data-stu-id="9bf81-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="9bf81-p102">アプリがアイテムの読み取りまたは編集フォームで表示されるかどうかを指定します。`Read`、`Edit` または `ReadOrEdit` のいずれかになります。`ExtensionPoint` 内の `Rule` で指定されている場合、この値は `Read` である必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="9bf81-136">**itemClass**</span><span class="sxs-lookup"><span data-stu-id="9bf81-136">**ItemClass**</span></span> | <span data-ttu-id="9bf81-137">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-137">No</span></span> | <span data-ttu-id="9bf81-p103">照合するカスタム メッセージ クラスを指定します。詳細については、「[特定のメッセージ クラスに対して Outlook のメール アドインをアクティブにする](https://docs.microsoft.com/outlook/add-ins/activation-rules)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="9bf81-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="9bf81-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="9bf81-141">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-141">No</span></span> | <span data-ttu-id="9bf81-142">アイテムが指定したメッセージ クラスのサブクラスである場合に、このルールは true と評価する必要があるかどうかを指定します。既定値は `false` です。</span><span class="sxs-lookup"><span data-stu-id="9bf81-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="9bf81-143">例</span><span class="sxs-lookup"><span data-stu-id="9bf81-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="9bf81-144">ItemHasAttachment ルール</span><span class="sxs-lookup"><span data-stu-id="9bf81-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="9bf81-145">アイテムに添付ファイルがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="9bf81-146">例</span><span class="sxs-lookup"><span data-stu-id="9bf81-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="9bf81-147">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="9bf81-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="9bf81-148">指定したエンティティ型のテキストがアイテムの件名または本文に含まれている場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="9bf81-149">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-149">Attributes</span></span>

| <span data-ttu-id="9bf81-150">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-150">Attribute</span></span> | <span data-ttu-id="9bf81-151">必須</span><span class="sxs-lookup"><span data-stu-id="9bf81-151">Required</span></span> | <span data-ttu-id="9bf81-152">説明</span><span class="sxs-lookup"><span data-stu-id="9bf81-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="9bf81-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="9bf81-153">**EntityType**</span></span> | <span data-ttu-id="9bf81-154">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-154">Yes</span></span> | <span data-ttu-id="9bf81-p104">このルールが true と評価するために見つける必要のあるエンティティの型を指定します。`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress`、または `Contact` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="9bf81-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="9bf81-157">**RegExFilter**</span></span> | <span data-ttu-id="9bf81-158">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-158">No</span></span> | <span data-ttu-id="9bf81-159">このエンティティに対してアクティブ化を実行するための正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="9bf81-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="9bf81-160">**FilterName**</span></span> | <span data-ttu-id="9bf81-161">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-161">No</span></span> | <span data-ttu-id="9bf81-162">正規表現フィルターの名前を指定します。指定すると、以後このフィルターをアドインのコード内で参照できます。</span><span class="sxs-lookup"><span data-stu-id="9bf81-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="9bf81-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="9bf81-163">**IgnoreCase**</span></span> | <span data-ttu-id="9bf81-164">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-164">No</span></span> | <span data-ttu-id="9bf81-165">**RegExFilter** 属性で指定した正規表現の実行時に、大文字と小文字の違いを無視するように指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="9bf81-166">**ハイライト**</span><span class="sxs-lookup"><span data-stu-id="9bf81-166">**Highlight**</span></span> | <span data-ttu-id="9bf81-167">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-167">No</span></span> | <span data-ttu-id="9bf81-p105">**注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するエンティティを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="9bf81-172">例</span><span class="sxs-lookup"><span data-stu-id="9bf81-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="9bf81-173">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="9bf81-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="9bf81-174">アイテムの指定したプロパティの中を検索し、指定した正規表現と一致するものがある場合に true と評価するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="9bf81-175">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-175">Attributes</span></span>

| <span data-ttu-id="9bf81-176">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-176">Attribute</span></span> | <span data-ttu-id="9bf81-177">必須</span><span class="sxs-lookup"><span data-stu-id="9bf81-177">Required</span></span> | <span data-ttu-id="9bf81-178">説明</span><span class="sxs-lookup"><span data-stu-id="9bf81-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="9bf81-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="9bf81-179">**RegExName**</span></span> | <span data-ttu-id="9bf81-180">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-180">Yes</span></span> | <span data-ttu-id="9bf81-181">アドインのコードで参照できるように、正規表現の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="9bf81-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="9bf81-182">**RegExValue**</span></span> | <span data-ttu-id="9bf81-183">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-183">Yes</span></span> | <span data-ttu-id="9bf81-184">メール アドインを表示するかどうかを判断するために評価する正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="9bf81-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="9bf81-185">**PropertyName**</span></span> | <span data-ttu-id="9bf81-186">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-186">Yes</span></span> | <span data-ttu-id="9bf81-p106">正規表現の評価対象となるプロパティの名前を指定します。`Subject`、`BodyAsPlaintext`、`BodyAsHtml`、または `SenderSTMPAddress` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="9bf81-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="9bf81-189">**IgnoreCase**</span></span> | <span data-ttu-id="9bf81-190">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-190">No</span></span> | <span data-ttu-id="9bf81-191">正規表現の実行時に大文字と小文字の違いを無視するように指定します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="9bf81-192">**ハイライト**</span><span class="sxs-lookup"><span data-stu-id="9bf81-192">**Highlight**</span></span> | <span data-ttu-id="9bf81-193">いいえ</span><span class="sxs-lookup"><span data-stu-id="9bf81-193">No</span></span> | <span data-ttu-id="9bf81-p107">**注意:** これは、**ExtensionPoint** 要素内の **Rule** 要素にのみ適用されます。クライアントが一致するテキストを強調表示にする方法を指定します。`all` または `none` のいずれかになります。指定のない場合、既定値は `all` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="9bf81-198">例</span><span class="sxs-lookup"><span data-stu-id="9bf81-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="9bf81-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="9bf81-199">RuleCollection</span></span>

<span data-ttu-id="9bf81-200">ルールのコレクション、およびそれらのルールの評価時に使用する論理演算子を定義します。</span><span class="sxs-lookup"><span data-stu-id="9bf81-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="9bf81-201">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-201">Attributes</span></span>

| <span data-ttu-id="9bf81-202">属性</span><span class="sxs-lookup"><span data-stu-id="9bf81-202">Attribute</span></span> | <span data-ttu-id="9bf81-203">必須</span><span class="sxs-lookup"><span data-stu-id="9bf81-203">Required</span></span> | <span data-ttu-id="9bf81-204">説明</span><span class="sxs-lookup"><span data-stu-id="9bf81-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="9bf81-205">**モード**</span><span class="sxs-lookup"><span data-stu-id="9bf81-205">**Mode**</span></span> | <span data-ttu-id="9bf81-206">はい</span><span class="sxs-lookup"><span data-stu-id="9bf81-206">Yes</span></span> | <span data-ttu-id="9bf81-p108">このルール コレクションの評価時に使用する論理演算子を指定します。`And` または `Or` のどちらかになります。</span><span class="sxs-lookup"><span data-stu-id="9bf81-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="9bf81-209">例</span><span class="sxs-lookup"><span data-stu-id="9bf81-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="9bf81-210">関連項目</span><span class="sxs-lookup"><span data-stu-id="9bf81-210">See also</span></span>

- [<span data-ttu-id="9bf81-211">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="9bf81-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="9bf81-212">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="9bf81-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="9bf81-213">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="9bf81-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)