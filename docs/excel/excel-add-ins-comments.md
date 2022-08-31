---
title: Excel JavaScript API を使用してコメントを操作する
description: API を使用してコメントスレッドとコメント スレッドを追加、削除、編集する方法に関する情報。
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5996c1bb55c3d4a358786b15f7c3e46aae6f42aa
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464798"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してコメントを操作する

この記事では、Excel JavaScript API を使用してブック内のコメントを追加、読み取り、変更、削除する方法について説明します。 コメント機能の詳細については、Excel の記事の [コメントとノートの挿入に関する記事を](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) 参照してください。

Excel JavaScript API では、コメントには、1 つの最初のコメントと接続されたスレッド化されたディスカッションの両方が含まれます。 これは個々のセルに関連付けられています。 十分なアクセス許可を持つブックを表示するすべてのユーザーは、コメントに返信できます。 [Comment](/javascript/api/excel/excel.comment) オブジェクトは、それらの応答を [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトとして格納します。 コメントをスレッドと見なす必要があります。スレッドには開始点として特別なエントリが必要です。

!["Comment.replies[0]" と "Comment.replies[1]" というラベルが付いた Excel コメント。](../images/excel-comments.png)

ブック内のコメントは、プロパティによって `Workbook.comments` 追跡されます。 これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。 `Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。 コメントには [ワークシート](/javascript/api/excel/excel.worksheet) レベルでもアクセスできます。 この記事のサンプルは、ブック レベルでコメントを操作しますが、プロパティを使用するように簡単に `Worksheet.comments` 変更できます。

## <a name="add-comments"></a>コメントを追加する

ブックに `CommentCollection.add` コメントを追加するには、このメソッドを使用します。 このメソッドは、最大 3 つのパラメーターを受け取ります。

- `cellAddress`: コメントが追加されるセル。 文字列オブジェクトまたは [Range](/javascript/api/excel/excel.range) オブジェクトを指定できます。 範囲は 1 つのセルである必要があります。
- `content`: コメントの内容。 プレーンテキスト コメントには文字列を使用します。 [メンション](#mentions)を含む[コメントには CommentRichContent](/javascript/api/excel/excel.commentrichcontent) オブジェクトを使用します。
- `contentType`: コンテンツの種類を指定する [ContentType](/javascript/api/excel/excel.contenttype) 列挙型。 既定値は `ContentType.plain` です。

次のコード例は、コメントをセル **A2** に追加します。

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> アドインによって追加されたコメントは、そのアドインの現在のユーザーに帰属します。

### <a name="add-comment-replies"></a>コメントの返信を追加する

`Comment`オブジェクトは、0 個以上の返信を含むコメント スレッドです。 `Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。 コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。 返信は、追加された順に表示されます。 また、アドインの現在のユーザーにも属性が設定されます。

次のコード サンプルは、ブックの最初のコメントに返信を追加します。

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## <a name="edit-comments"></a>コメントを編集する

コメントまたはコメントの返信を編集するには、その `Comment.content` プロパティまたは `CommentReply.content` プロパティを設定します。

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### <a name="edit-comment-replies"></a>コメントの返信を編集する

コメントの返信を編集するには、そのプロパティを `CommentReply.content` 設定します。

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## <a name="delete-comments"></a>コメントを削除する

コメントを削除するには、メソッドを使用します `Comment.delete` 。 コメントを削除すると、そのコメントに関連付けられている返信も削除されます。

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### <a name="delete-comment-replies"></a>コメントの返信を削除する

コメントの返信を削除するには、メソッドを使用します `CommentReply.delete` 。

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="resolve-comment-threads"></a>コメント スレッドを解決する

コメント スレッドには、 `resolved`解決されたかどうかを示す、構成可能なブール値があります。 値は `true` 、コメント スレッドが解決されたことを意味します。 値 `false` は、コメント スレッドが新しいか再び開かれたかのどちらかであることを意味します。

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

コメントの返信には読み取り専用 `resolved` プロパティがあります。 その値は、常にスレッドの残りの部分の値と等しくなります。

## <a name="comment-metadata"></a>コメント メタデータ

各コメントには、作成者や作成日などの作成に関するメタデータが含まれています。 アドインによって作成されたコメントは、現在のユーザーによって作成されたものと見なされます。

次のサンプルは、**A2** に作成者のメール、作成者の名前、コメントの作成日を表示する方法を示しています。

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### <a name="comment-reply-metadata"></a>コメント応答メタデータ

コメントの返信には、最初のコメントと同じ種類のメタデータが格納されます。

次の例では、 **A2** で最新のコメント返信の作成者のメール、作成者の名前、作成日を表示する方法を示します。

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## <a name="mentions"></a>メンション

[メンション](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) は、コメント内の同僚にタグを付けるために使用されます。 これにより、コメントの内容を含む通知が送信されます。 アドインは、ユーザーに代わってこれらのメンションを作成できます。

メンションを含むコメントは、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) オブジェクトを使用して作成する必要があります。 1 つ以上の`CommentRichContent`メンションを含む呼び出`CommentCollection.add`しで、パラメーターとして`contentType`指定`ContentType.mention`します。 `content`テキストにメンションを挿入するには、文字列の書式設定も必要です。 メンションの形式は次のとおりです `<at id="{replyIndex}">{mentionName}</at>`。

> [!NOTE]
> 現在、メンション リンクのテキストとして使用できるのは、メンションの正確な名前のみです。 名前の短縮バージョンのサポートは、後で追加される予定です。

次の例は、1 つのメンションを含むコメントを示しています。

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## <a name="comment-events"></a>コメント イベント

アドインは、コメントの追加、変更、削除をリッスンできます。 [コメント イベント](/javascript/api/excel/excel.commentcollection#event-details) は、オブジェクトに対して `CommentCollection` 発生します。 コメント イベントをリッスンするには、イベント ハンドラー 、または`onDeleted`コメント イベント ハンドラーを`onAdded``onChanged`登録します。 コメント イベントが検出されたら、このイベント ハンドラーを使用して、追加、変更、または削除されたコメントに関するデータを取得します。 このイベントでは `onChanged` 、コメント応答の追加、変更、削除も処理されます。

各コメント イベントは、複数の追加、変更、または削除が同時に実行されたときに 1 回だけトリガーされます。 [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)、[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) のすべてのオブジェクトには、イベント アクションをコメント コレクションにマップするためのコメント ID の配列が含まれています。

イベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の詳細については、 [Excel JavaScript API を使用](excel-add-ins-events.md) したイベントの処理に関する記事を参照してください。

### <a name="comment-addition-events"></a>コメント追加イベント

イベントは `onAdded` 、1 つ以上の新しいコメントがコメント コレクションに追加されたときにトリガーされます。 このイベントは、返信がコメント スレッドに追加されたときにトリガー *されません* ( [コメントの応答イベント](#comment-change-events) について詳しくは、「コメント変更イベント」をご覧ください)。

次の例では、イベント ハンドラーを `onAdded` 登録し、オブジェクトを `CommentAddedEventArgs` 使用して追加されたコメントの配列を `commentDetails` 取得する方法を示します。

> [!NOTE]
> このサンプルは、1 つのコメントが追加された場合にのみ機能します。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-change-events"></a>コメント変更イベント

`onChanged`コメント イベントは、次のシナリオでトリガーされます。

- コメントの内容が更新されます。
- コメント スレッドが解決されます。
- コメント スレッドが再び開きます。
- コメント スレッドに返信が追加されます。
- コメント スレッドで返信が更新されます。
- コメント スレッドで応答が削除されます。

次の例では、イベント ハンドラーを `onChanged` 登録し、オブジェクトを `CommentChangedEventArgs` 使用して変更されたコメントの配列を `commentDetails` 取得する方法を示します。

> [!NOTE]
> このサンプルは、1 つのコメントが変更された場合にのみ機能します。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-deletion-events"></a>コメント削除イベント

イベントは `onDeleted` 、コメントがコメント コレクションから削除されたときにトリガーされます。 コメントが削除されると、そのメタデータは使用できなくなります。 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) オブジェクトは、アドインが個々のコメントを管理している場合に備えて、コメント ID を提供します。

次の例では、イベント ハンドラーを `onDeleted` 登録し、オブジェクトを `CommentDeletedEventArgs` 使用して削除されたコメントの配列を取得 `commentDetails` する方法を示します。

> [!NOTE]
> このサンプルは、1 つのコメントが削除された場合にのみ機能します。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してブックを操作する](excel-add-ins-workbooks.md)
- [Excel JavaScript API を使用してイベントを操作する](excel-add-ins-events.md)
- [Excel でコメントとメモを挿入する](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
