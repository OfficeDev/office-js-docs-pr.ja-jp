---
title: Excel JavaScript API を使用してコメントを操作する
description: Api を使用してコメントおよびコメントスレッドを追加、削除、および編集する方法について説明します。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 85312cbd92aa6c9d0f82fd167e8a372c2eff8c85
ms.sourcegitcommit: b50eebd303adcc22eb86e65756ce7e9a82f41a57
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/14/2020
ms.locfileid: "48456553"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="18c44-103">Excel JavaScript API を使用してコメントを操作する</span><span class="sxs-lookup"><span data-stu-id="18c44-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="18c44-104">この記事では、Excel JavaScript API を使用してブック内のコメントを追加、読み取り、変更、および削除する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="18c44-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="18c44-105">コメント機能の詳細については、「 [Excel 記事のコメントとメモを挿入する」](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="18c44-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="18c44-106">Excel JavaScript API では、コメントには単一の最初のコメントと接続されたスレッドのディスカッションの両方が含まれます。</span><span class="sxs-lookup"><span data-stu-id="18c44-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="18c44-107">個別のセルに関連付けられています。</span><span class="sxs-lookup"><span data-stu-id="18c44-107">It is tied to an individual cell.</span></span> <span data-ttu-id="18c44-108">十分な権限があるブックを表示するユーザーは、コメントに返信できます。</span><span class="sxs-lookup"><span data-stu-id="18c44-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="18c44-109">Comment オブジェクトは、これらの返信を[コメント](/javascript/api/excel/excel.comment)[返信](/javascript/api/excel/excel.commentreply)オブジェクトとして格納します。</span><span class="sxs-lookup"><span data-stu-id="18c44-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="18c44-110">コメントはスレッドと考えてください。スレッドには、開始点として特別なエントリが必要です。</span><span class="sxs-lookup"><span data-stu-id="18c44-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![「Comment」というラベルが付けられた、"comment" というラベルが付いた Excel コメント。「comment [0]」と「Comment [1]」。](../images/excel-comments.png)

<span data-ttu-id="18c44-112">ブック内のコメントはプロパティによって追跡され `Workbook.comments` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="18c44-113">これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。</span><span class="sxs-lookup"><span data-stu-id="18c44-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="18c44-114">`Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="18c44-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="18c44-115">コメントには、 [ワークシート](/javascript/api/excel/excel.worksheet) レベルでアクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="18c44-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="18c44-116">この記事のサンプルでは、ブックレベルでコメントを使用していますが、プロパティを使用するために簡単に変更することができ `Worksheet.comments` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="18c44-117">コメントを追加する</span><span class="sxs-lookup"><span data-stu-id="18c44-117">Add comments</span></span>

<span data-ttu-id="18c44-118">メソッドを使用して、 `CommentCollection.add` ブックにコメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="18c44-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="18c44-119">このメソッドは、次の3つのパラメーターを取ります。</span><span class="sxs-lookup"><span data-stu-id="18c44-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="18c44-120">`cellAddress`: コメントが追加されるセルを指定します。</span><span class="sxs-lookup"><span data-stu-id="18c44-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="18c44-121">文字列または [Range](/javascript/api/excel/excel.range) オブジェクトのいずれかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="18c44-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="18c44-122">範囲は1つのセルである必要があります。</span><span class="sxs-lookup"><span data-stu-id="18c44-122">The range must be a single cell.</span></span>
- <span data-ttu-id="18c44-123">`content`: コメントの内容。</span><span class="sxs-lookup"><span data-stu-id="18c44-123">`content`: The comment's content.</span></span> <span data-ttu-id="18c44-124">テキスト形式のコメントには文字列を使用します。</span><span class="sxs-lookup"><span data-stu-id="18c44-124">Use a string for plain text comments.</span></span> <span data-ttu-id="18c44-125">[メンション](#mentions)付きのコメントには、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="18c44-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="18c44-126">`contentType`: コンテンツの種類を指定する [ContentType](/javascript/api/excel/excel.contenttype) 列挙。</span><span class="sxs-lookup"><span data-stu-id="18c44-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="18c44-127">既定値は `ContentType.plain` です。</span><span class="sxs-lookup"><span data-stu-id="18c44-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="18c44-128">次のコード例は、コメントをセル **A2** に追加します。</span><span class="sxs-lookup"><span data-stu-id="18c44-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="18c44-129">アドインによって追加されたコメントは、そのアドインの現在のユーザーによって作成されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="18c44-130">コメントの返信を追加する</span><span class="sxs-lookup"><span data-stu-id="18c44-130">Add comment replies</span></span>

<span data-ttu-id="18c44-131">`Comment`オブジェクトは、0個以上の返信を含むコメントスレッドです。</span><span class="sxs-lookup"><span data-stu-id="18c44-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="18c44-132">`Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。</span><span class="sxs-lookup"><span data-stu-id="18c44-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="18c44-133">コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。</span><span class="sxs-lookup"><span data-stu-id="18c44-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="18c44-134">返信は、追加された順に表示されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="18c44-135">また、アドインの現在のユーザーにも属性があります。</span><span class="sxs-lookup"><span data-stu-id="18c44-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="18c44-136">次のコード サンプルは、ブックの最初のコメントに返信を追加します。</span><span class="sxs-lookup"><span data-stu-id="18c44-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="18c44-137">コメントの編集</span><span class="sxs-lookup"><span data-stu-id="18c44-137">Edit comments</span></span>

<span data-ttu-id="18c44-138">コメントまたはコメントの返信を編集するには、その `Comment.content` プロパティまたは `CommentReply.content` プロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="18c44-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="18c44-139">コメントの返信を編集する</span><span class="sxs-lookup"><span data-stu-id="18c44-139">Edit comment replies</span></span>

<span data-ttu-id="18c44-140">コメントの返信を編集するには、そのプロパティを設定 `CommentReply.content` します。</span><span class="sxs-lookup"><span data-stu-id="18c44-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="18c44-141">コメントの削除</span><span class="sxs-lookup"><span data-stu-id="18c44-141">Delete comments</span></span>

<span data-ttu-id="18c44-142">コメントを削除するには、メソッドを使用し `Comment.delete` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="18c44-143">コメントを削除すると、そのコメントに関連付けられている返信も削除されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="18c44-144">コメントの返信を削除する</span><span class="sxs-lookup"><span data-stu-id="18c44-144">Delete comment replies</span></span>

<span data-ttu-id="18c44-145">コメントの返信を削除するには、メソッドを使用し `CommentReply.delete` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="18c44-146">コメントスレッドを解決する</span><span class="sxs-lookup"><span data-stu-id="18c44-146">Resolve comment threads</span></span>

<span data-ttu-id="18c44-147">コメントスレッドには、解決可能かどうかを示す、構成可能なブール値があり `resolved` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="18c44-148">の値は、 `true` コメントスレッドが解決されたことを意味します。</span><span class="sxs-lookup"><span data-stu-id="18c44-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="18c44-149">の値は、 `false` コメントスレッドが新規または再オープンのいずれかであることを意味します。</span><span class="sxs-lookup"><span data-stu-id="18c44-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="18c44-150">コメントの返信には、readonly プロパティがあり `resolved` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="18c44-151">この値は、常にスレッドの残りの部分と同じです。</span><span class="sxs-lookup"><span data-stu-id="18c44-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="18c44-152">コメントのメタデータ</span><span class="sxs-lookup"><span data-stu-id="18c44-152">Comment metadata</span></span>

<span data-ttu-id="18c44-153">各コメントには、作成者や作成日などの作成に関するメタデータが含まれています。</span><span class="sxs-lookup"><span data-stu-id="18c44-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="18c44-154">アドインによって作成されたコメントは、現在のユーザーによって作成されたものと見なされます。</span><span class="sxs-lookup"><span data-stu-id="18c44-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="18c44-155">次のサンプルは、**A2** に作成者のメール、作成者の名前、コメントの作成日を表示する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="18c44-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="18c44-156">コメントの返信メタデータ</span><span class="sxs-lookup"><span data-stu-id="18c44-156">Comment reply metadata</span></span>

<span data-ttu-id="18c44-157">コメントの返信は、最初のコメントと同じ種類のメタデータを格納します。</span><span class="sxs-lookup"><span data-stu-id="18c44-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="18c44-158">次の例は、作成者の電子メール、作成者の名前、および **A2**における最新のコメントの返信の作成日を表示する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="18c44-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="18c44-159">メンション</span><span class="sxs-lookup"><span data-stu-id="18c44-159">Mentions</span></span>

<span data-ttu-id="18c44-160">[メンション](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) は、コメント内の仕事仲間にタグ付けするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="18c44-161">これにより、それらの通知がコメントの内容と共に送信されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="18c44-162">アドインは、ユーザーの代わりにこれらのメンションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="18c44-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="18c44-163">[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用して、メンションを含むコメントを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="18c44-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="18c44-164">1つ以上のメンションを含むを呼び出し、 `CommentCollection.add` `CommentRichContent` `ContentType.mention` パラメーターとしてを指定し `contentType` ます。</span><span class="sxs-lookup"><span data-stu-id="18c44-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="18c44-165">`content`文字列をテキストに挿入するには、文字列を書式設定する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="18c44-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="18c44-166">メンションの形式は、 `<at id="{replyIndex}">{mentionName}</at>` です。</span><span class="sxs-lookup"><span data-stu-id="18c44-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="18c44-167">現時点では、メンションリンクのテキストとして、メンションの正確な名前のみを使用できます。</span><span class="sxs-lookup"><span data-stu-id="18c44-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="18c44-168">名前の短縮バージョンのサポートは、後で追加されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="18c44-169">次の例は、1つのメンション付きのコメントを示しています。</span><span class="sxs-lookup"><span data-stu-id="18c44-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="comment-events"></a><span data-ttu-id="18c44-170">コメントイベント</span><span class="sxs-lookup"><span data-stu-id="18c44-170">Comment events</span></span>

<span data-ttu-id="18c44-171">アドインは、コメントの追加、変更、および削除を聞くことができます。</span><span class="sxs-lookup"><span data-stu-id="18c44-171">Your add-in can listen for comment additions, changes, and deletions.</span></span> <span data-ttu-id="18c44-172">[Comment イベント](/javascript/api/excel/excel.commentcollection#event-details) は、オブジェクトに対して発生 `CommentCollection` します。</span><span class="sxs-lookup"><span data-stu-id="18c44-172">[Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.</span></span> <span data-ttu-id="18c44-173">Comment イベントをリッスンするには、、、 `onAdded` `onChanged` またはの `onDeleted` コメントイベントハンドラーを登録します。</span><span class="sxs-lookup"><span data-stu-id="18c44-173">To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler.</span></span> <span data-ttu-id="18c44-174">コメントイベントが検出されたときに、追加、変更、または削除されたコメントに関するデータを取得するには、このイベントハンドラーを使用します。</span><span class="sxs-lookup"><span data-stu-id="18c44-174">When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment.</span></span> <span data-ttu-id="18c44-175">この `onChanged` イベントは、コメントの返信の追加、変更、および削除も処理します。</span><span class="sxs-lookup"><span data-stu-id="18c44-175">The `onChanged` event also handles comment reply additions, changes, and deletions.</span></span> 

<span data-ttu-id="18c44-176">各 comment イベントは、同時に複数の追加、変更、または削除が実行された場合にのみトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="18c44-176">Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time.</span></span> <span data-ttu-id="18c44-177">[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、 [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg)、および[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)のすべてのオブジェクトには、イベントアクションをコメントのコレクションにマップするためのコメント id の配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="18c44-177">All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.</span></span>

<span data-ttu-id="18c44-178">イベントハンドラーの登録、イベントの処理、イベントハンドラーの削除に関する追加情報については、「 [Excel JAVASCRIPT API を使用してイベント](excel-add-ins-events.md) を処理する」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="18c44-178">See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.</span></span> 

### <a name="comment-addition-events"></a><span data-ttu-id="18c44-179">コメントの追加イベント</span><span class="sxs-lookup"><span data-stu-id="18c44-179">Comment addition events</span></span> 
<span data-ttu-id="18c44-180">この `onAdded` イベントは、コメントのコレクションに1つまたは複数の新しいコメントが追加されると発生します。</span><span class="sxs-lookup"><span data-stu-id="18c44-180">The `onAdded` event is triggered when one or more new comments are added to the comment collection.</span></span> <span data-ttu-id="18c44-181">このイベントは、コメントスレッドに返信が追加されたときには発生し *ません* (コメントの返信イベントについては、「 [コメント変更イベント](#comment-change-events) 」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="18c44-181">This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).</span></span>

<span data-ttu-id="18c44-182">次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して追加されたコメントの配列を取得する方法を示して `onAdded` `CommentAddedEventArgs` `commentDetails` います。</span><span class="sxs-lookup"><span data-stu-id="18c44-182">The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.</span></span>

> [!NOTE]
> <span data-ttu-id="18c44-183">このサンプルは、1つのコメントが追加された場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="18c44-183">This sample only works when a single comment is added.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a><span data-ttu-id="18c44-184">コメント変更イベント</span><span class="sxs-lookup"><span data-stu-id="18c44-184">Comment change events</span></span> 
<span data-ttu-id="18c44-185">`onChanged`Comment イベントは、次のシナリオでトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="18c44-185">The `onChanged` comment event is triggered in the following scenarios.</span></span>

- <span data-ttu-id="18c44-186">コメントの内容が更新されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-186">A comment's content is updated.</span></span>
- <span data-ttu-id="18c44-187">コメントスレッドが解決されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-187">A comment thread is resolved.</span></span>
- <span data-ttu-id="18c44-188">コメントスレッドが再度開かれています。</span><span class="sxs-lookup"><span data-stu-id="18c44-188">A comment thread is reopened.</span></span>
- <span data-ttu-id="18c44-189">コメントスレッドに返信が追加されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-189">A reply is added to a comment thread.</span></span>
- <span data-ttu-id="18c44-190">コメントスレッド内の返信が更新されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-190">A reply is updated in a comment thread.</span></span>
- <span data-ttu-id="18c44-191">コメントスレッド内の返信が削除されます。</span><span class="sxs-lookup"><span data-stu-id="18c44-191">A reply is deleted in a comment thread.</span></span>

<span data-ttu-id="18c44-192">次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して、変更されたコメントの配列を取得する方法を示して `onChanged` `CommentChangedEventArgs` `commentDetails` います。</span><span class="sxs-lookup"><span data-stu-id="18c44-192">The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.</span></span>

> [!NOTE]
> <span data-ttu-id="18c44-193">このサンプルは、1つのコメントが変更された場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="18c44-193">This sample only works when a single comment is changed.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a><span data-ttu-id="18c44-194">コメント削除イベント</span><span class="sxs-lookup"><span data-stu-id="18c44-194">Comment deletion events</span></span>
<span data-ttu-id="18c44-195">コメントの `onDeleted` コレクションからコメントが削除されると、イベントがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="18c44-195">The `onDeleted` event is triggered when a comment is deleted from the comment collection.</span></span> <span data-ttu-id="18c44-196">コメントが削除されると、そのメタデータは使用できなくなります。</span><span class="sxs-lookup"><span data-stu-id="18c44-196">Once a comment has been deleted, its metadata is no longer available.</span></span> <span data-ttu-id="18c44-197">[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)オブジェクトは、アドインが個々のコメントを管理している場合に、コメント id を提供します。</span><span class="sxs-lookup"><span data-stu-id="18c44-197">The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.</span></span>

<span data-ttu-id="18c44-198">次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して、削除されたコメントの配列を取得する方法を示して `onDeleted` `CommentDeletedEventArgs` `commentDetails` います。</span><span class="sxs-lookup"><span data-stu-id="18c44-198">The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.</span></span>

> [!NOTE]
> <span data-ttu-id="18c44-199">このサンプルは、1つのコメントが削除された場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="18c44-199">This sample only works when a single comment is deleted.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a><span data-ttu-id="18c44-200">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="18c44-200">See also</span></span>

- [<span data-ttu-id="18c44-201">Office アドインでの Excel JavaScript オブジェクトモデル</span><span class="sxs-lookup"><span data-stu-id="18c44-201">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="18c44-202">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="18c44-202">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="18c44-203">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="18c44-203">Work with Events using the Excel JavaScript API</span></span>](excel-add-ins-events.md)
- [<span data-ttu-id="18c44-204">Excel でコメントやメモを挿入する</span><span class="sxs-lookup"><span data-stu-id="18c44-204">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
