---
title: Excel JavaScript API を使用してコメントを操作する
description: ''
ms.date: 10/22/2019
localization_priority: Normal
ms.openlocfilehash: d79f99d1922def58fe2c8887d01ec5a2b173220a
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37681915"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="8cce0-102">Excel JavaScript API を使用してコメントを操作する</span><span class="sxs-lookup"><span data-stu-id="8cce0-102">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="8cce0-103">この記事では、Excel JavaScript API を使用してブック内のコメントを追加、読み取り、変更、および削除する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-103">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="8cce0-104">コメント機能の詳細については、「 [Excel 記事のコメントとメモを挿入する」](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8cce0-104">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="8cce0-105">Excel JavaScript API では、コメントは最初のメモと接続されたスレッドのディスカッションの両方です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-105">In the Excel JavaScript API, a comment is both the initial note and the connected threaded discussion.</span></span> <span data-ttu-id="8cce0-106">個別のセルに関連付けられています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-106">It is tied to an individual cell.</span></span> <span data-ttu-id="8cce0-107">十分な権限があるブックを表示するユーザーは、コメントに返信できます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-107">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="8cce0-108">Comment オブジェクトは、これらの返信を[コメント](/javascript/api/excel/excel.comment)[返信](/javascript/api/excel/excel.commentreply)オブジェクトとして格納します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-108">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="8cce0-109">コメントはスレッドと考えてください。スレッドには、開始点として特別なエントリが必要です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-109">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![「Comment」というラベルが付けられた、"comment" というラベルが付いた Excel コメント。「comment [0]」と「Comment [1]」。](../images/excel-comments.png)

<span data-ttu-id="8cce0-111">ブック内のコメントは`Workbook.comments`プロパティによって追跡されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-111">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="8cce0-112">これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-112">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="8cce0-113">`Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8cce0-113">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="8cce0-114">コメントには、[ワークシート](/javascript/api/excel/excel.worksheet)レベルでアクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-114">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="8cce0-115">この記事のサンプルでは、ブックレベルでコメントを使用していますが、 `Worksheet.comments`プロパティを使用するために簡単に変更することができます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-115">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="8cce0-116">コメントを追加する</span><span class="sxs-lookup"><span data-stu-id="8cce0-116">Add comments</span></span>

<span data-ttu-id="8cce0-117">メソッドを`CommentCollection.add`使用して、ブックにコメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-117">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="8cce0-118">このメソッドは、次の3つのパラメーターを取ります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-118">This method takes up to three parameters:</span></span>

- <span data-ttu-id="8cce0-119">`cellAddress`: コメントが追加されるセルを指定します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-119">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="8cce0-120">文字列または[Range](/javascript/api/excel/excel.range)オブジェクトのいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-120">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="8cce0-121">範囲は1つのセルである必要があります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-121">The range must be a single cell.</span></span>
- <span data-ttu-id="8cce0-122">`content`: コメントの内容。</span><span class="sxs-lookup"><span data-stu-id="8cce0-122">`content`: The comment's content.</span></span> <span data-ttu-id="8cce0-123">テキスト形式のコメントには文字列を使用します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-123">Use a string for plain text comments.</span></span> <span data-ttu-id="8cce0-124">[メンション](#mentions-preview)付きのコメントには、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-124">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions-preview).</span></span>
- <span data-ttu-id="8cce0-125">`contentType`: コンテンツの種類を指定する[ContentType](/javascript/api/excel/excel.contenttype)列挙。</span><span class="sxs-lookup"><span data-stu-id="8cce0-125">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="8cce0-126">既定値は `ContentType.plain` です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-126">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="8cce0-127">次のコード例は、コメントをセル **A2** に追加します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-127">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> <span data-ttu-id="8cce0-128">アドインによって追加されたコメントは、そのアドインの現在のユーザーによって作成されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-128">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="8cce0-129">コメントの返信を追加する</span><span class="sxs-lookup"><span data-stu-id="8cce0-129">Add comment replies</span></span>

<span data-ttu-id="8cce0-130">`Comment`オブジェクトは、0個以上の返信を含むコメントスレッドです。</span><span class="sxs-lookup"><span data-stu-id="8cce0-130">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="8cce0-131">`Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-131">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="8cce0-132">コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-132">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="8cce0-133">返信は、追加された順に表示されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-133">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="8cce0-134">また、アドインの現在のユーザーにも属性があります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-134">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="8cce0-135">次のコード サンプルは、ブックの最初のコメントに返信を追加します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-135">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="8cce0-136">コメントの編集</span><span class="sxs-lookup"><span data-stu-id="8cce0-136">Edit comments</span></span>

<span data-ttu-id="8cce0-137">コメントまたはコメントの返信を編集するには、その `Comment.content` プロパティまたは `CommentReply.content` プロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-137">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="8cce0-138">コメントの返信を編集する</span><span class="sxs-lookup"><span data-stu-id="8cce0-138">Edit comment replies</span></span>

<span data-ttu-id="8cce0-139">コメントの返信を編集するには`CommentReply.content` 、そのプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-139">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="8cce0-140">コメントの削除</span><span class="sxs-lookup"><span data-stu-id="8cce0-140">Delete comments</span></span>

<span data-ttu-id="8cce0-141">コメントを削除するには`Comment.delete` 、メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-141">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="8cce0-142">コメントを削除すると、そのコメントに関連付けられている返信も削除されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-142">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="8cce0-143">コメントの返信を削除する</span><span class="sxs-lookup"><span data-stu-id="8cce0-143">Delete comment replies</span></span>

<span data-ttu-id="8cce0-144">コメントの返信を削除するには`CommentReply.delete` 、メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-144">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="8cce0-145">コメントスレッドを解決する</span><span class="sxs-lookup"><span data-stu-id="8cce0-145">Resolve comment threads</span></span>

<span data-ttu-id="8cce0-146">コメントスレッドには、解決可能かどう`resolved`かを示す、構成可能なブール値があります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-146">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="8cce0-147">の`true`値は、コメントスレッドが解決されたことを意味します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-147">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="8cce0-148">の`false`値は、コメントスレッドが新規または再オープンのいずれかであることを意味します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-148">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="8cce0-149">コメントの返信には`resolved` 、readonly プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-149">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="8cce0-150">この値は、常にスレッドの残りの部分と同じです。</span><span class="sxs-lookup"><span data-stu-id="8cce0-150">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="8cce0-151">コメントのメタデータ</span><span class="sxs-lookup"><span data-stu-id="8cce0-151">Comment metadata</span></span>

<span data-ttu-id="8cce0-152">各コメントには、作成者や作成日などの作成に関するメタデータが含まれています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-152">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="8cce0-153">アドインによって作成されたコメントは、現在のユーザーによって作成されたものと見なされます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-153">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="8cce0-154">次のサンプルは、**A2** に作成者のメール、作成者の名前、コメントの作成日を表示する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-154">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a><span data-ttu-id="8cce0-155">コメントの返信メタデータ</span><span class="sxs-lookup"><span data-stu-id="8cce0-155">Comment reply metadata</span></span>

<span data-ttu-id="8cce0-156">コメントの返信は、最初のコメントと同じ種類のメタデータを格納します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-156">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="8cce0-157">次の例は、作成者の電子メール、作成者の名前、および**A2**における最新のコメントの返信の作成日を表示する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-157">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions-preview"></a><span data-ttu-id="8cce0-158">メンション (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8cce0-158">Mentions (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="8cce0-159">コメントコメント Api は、現在、パブリックプレビューでのみ利用可能です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-159">The comment mention APIs are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> <span data-ttu-id="8cce0-160">コメントメンションは、現在 web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-160">Comment mentions are currently only supported for Excel on the web.</span></span>

<span data-ttu-id="8cce0-161">[メンション](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd)は、コメント内の仕事仲間にタグ付けするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-161">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="8cce0-162">これにより、それらの通知がコメントの内容と共に送信されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-162">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="8cce0-163">アドインは、ユーザーの代わりにこれらのメンションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-163">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="8cce0-164">[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用して、メンションを含むコメントを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-164">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="8cce0-165">1 `CommentCollection.add`つ以上`CommentRichContent`のメンションを含むを呼び出し`ContentType.mention` 、 `contentType`パラメーターとしてを指定します。</span><span class="sxs-lookup"><span data-stu-id="8cce0-165">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="8cce0-166">`content`文字列をテキストに挿入するには、文字列を書式設定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="8cce0-166">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="8cce0-167">メンションの形式は、 `<at id="{replyIndex}">{mentionName}</at>`です。</span><span class="sxs-lookup"><span data-stu-id="8cce0-167">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> <span data-ttu-id="8cce0-168">こと現時点では、メンションリンクのテキストとして、メンションの正確な名前のみを使用できます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-168">[NOTE] Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="8cce0-169">名前の短縮バージョンのサポートは、後で追加されます。</span><span class="sxs-lookup"><span data-stu-id="8cce0-169">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="8cce0-170">次の例は、1つのメンション付きのコメントを示しています。</span><span class="sxs-lookup"><span data-stu-id="8cce0-170">The following example shows a comment with a single mention.</span></span>

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="8cce0-171">関連項目</span><span class="sxs-lookup"><span data-stu-id="8cce0-171">See also</span></span>

- [<span data-ttu-id="8cce0-172">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="8cce0-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8cce0-173">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="8cce0-173">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="8cce0-174">Excel でコメントやメモを挿入する</span><span class="sxs-lookup"><span data-stu-id="8cce0-174">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)