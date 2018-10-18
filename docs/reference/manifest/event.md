# <a name="event-element"></a><span data-ttu-id="89669-101">Event 要素</span><span class="sxs-lookup"><span data-stu-id="89669-101">Event element</span></span>

<span data-ttu-id="89669-102">アドインでイベント ハンドラを定義します。</span><span class="sxs-lookup"><span data-stu-id="89669-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="89669-103">`Event` 要素は現在、Office 365 の Outlook on the web でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="89669-103">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="89669-104">属性</span><span class="sxs-lookup"><span data-stu-id="89669-104">Attributes</span></span>

|  <span data-ttu-id="89669-105">属性</span><span class="sxs-lookup"><span data-stu-id="89669-105">Attribute</span></span>  |  <span data-ttu-id="89669-106">必須</span><span class="sxs-lookup"><span data-stu-id="89669-106">Required</span></span>  |  <span data-ttu-id="89669-107">説明</span><span class="sxs-lookup"><span data-stu-id="89669-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="89669-108">型</span><span class="sxs-lookup"><span data-stu-id="89669-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="89669-109">はい</span><span class="sxs-lookup"><span data-stu-id="89669-109">Yes</span></span>  | <span data-ttu-id="89669-110">処理するイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="89669-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="89669-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="89669-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="89669-112">はい</span><span class="sxs-lookup"><span data-stu-id="89669-112">Yes</span></span>  | <span data-ttu-id="89669-p101">イベント ハンドラの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラのみです。</span><span class="sxs-lookup"><span data-stu-id="89669-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="89669-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="89669-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="89669-116">はい</span><span class="sxs-lookup"><span data-stu-id="89669-116">Yes</span></span>  | <span data-ttu-id="89669-117">イベント ハンドラの関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="89669-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="89669-118">Type 属性</span><span class="sxs-lookup"><span data-stu-id="89669-118">Type attribute</span></span>

<span data-ttu-id="89669-p102">必須です。イベント ハンドラを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。</span><span class="sxs-lookup"><span data-stu-id="89669-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="89669-122">イベントの種類</span><span class="sxs-lookup"><span data-stu-id="89669-122">Event type</span></span>  |  <span data-ttu-id="89669-123">説明</span><span class="sxs-lookup"><span data-stu-id="89669-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="89669-124">ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="89669-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="89669-125">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="89669-125">FunctionExecution attribute</span></span>

<span data-ttu-id="89669-126">必須。</span><span class="sxs-lookup"><span data-stu-id="89669-126">Required.</span></span> <span data-ttu-id="89669-127">`synchronous` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89669-127">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="89669-128">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="89669-128">FunctionName attribute</span></span>

<span data-ttu-id="89669-p104">必須です。イベント ハンドラの関数名を指定します。この値は、アドインの [ 関数ファイル](functionfile.md)内の関数名と一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89669-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```