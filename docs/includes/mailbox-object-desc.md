<span data-ttu-id="eca76-101">Outlook アドインは、主に[Mailbox](/javascript/api/outlook/office.mailbox)オブジェクトを介して公開される api を使用します。</span><span class="sxs-lookup"><span data-stu-id="eca76-101">Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/office.mailbox) object.</span></span> <span data-ttu-id="eca76-102">Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="eca76-102">To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.</span></span>

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

<span data-ttu-id="eca76-103">さらに、Outlook アドインでは次のオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="eca76-103">Additionally, Outlook add-ins can use the following objects:</span></span>

-  <span data-ttu-id="eca76-104">**Office** オブジェクト: 初期化に使用します。</span><span class="sxs-lookup"><span data-stu-id="eca76-104">**Office** object: for initialization.</span></span>

-  <span data-ttu-id="eca76-105">**Context** オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。</span><span class="sxs-lookup"><span data-stu-id="eca76-105">**Context** object: for access to content and display language properties.</span></span>

-  <span data-ttu-id="eca76-106">**RoamingSettings** オブジェクト: アドインがインストールされているユーザーのメールボックスに Outlook アドイン固有のカスタム設定を保存する際に使用します。</span><span class="sxs-lookup"><span data-stu-id="eca76-106">**RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.</span></span>

<span data-ttu-id="eca76-107">Outlook JavaScript API の使用方法については、「 [outlook アドイン](../outlook/outlook-add-ins-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eca76-107">For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).</span></span>