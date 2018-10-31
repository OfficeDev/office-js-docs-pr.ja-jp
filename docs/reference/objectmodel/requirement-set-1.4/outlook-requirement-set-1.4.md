# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="11e31-101">Outlook アドイン API 要件セット 1.4</span><span class="sxs-lookup"><span data-stu-id="11e31-101">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="11e31-102">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="11e31-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="11e31-103">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)が対象です。</span><span class="sxs-lookup"><span data-stu-id="11e31-103">Note: This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="11e31-104">1.4 の新機能</span><span class="sxs-lookup"><span data-stu-id="11e31-104">What's new in 1.4?</span></span>

<span data-ttu-id="11e31-p101">要件セット 1.4 は、[要件セット 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) のすべての機能を含み、`Office.ui` 名前空間へのアクセスが追加されています。</span><span class="sxs-lookup"><span data-stu-id="11e31-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="11e31-107">変更ログ</span><span class="sxs-lookup"><span data-stu-id="11e31-107">Change log</span></span>

- <span data-ttu-id="11e31-108">[Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) を追加: Office ホストでダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="11e31-108">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="11e31-109">[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) を追加: メッセージをダイアログ ボックスからその親/オープナー ページに配信します。</span><span class="sxs-lookup"><span data-stu-id="11e31-109">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="11e31-110">[Dialog](/javascript/api/office/office.dialog)  オブジェクトを追加:  [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) メソッドが呼び出されたときに返されるオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="11e31-110">Added Dialog object: The object that is returned when the  method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="11e31-111">関連項目</span><span class="sxs-lookup"><span data-stu-id="11e31-111">See also</span></span>

- [<span data-ttu-id="11e31-112">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="11e31-112">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="11e31-113">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="11e31-113">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="11e31-114">作業の開始</span><span class="sxs-lookup"><span data-stu-id="11e31-114">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)