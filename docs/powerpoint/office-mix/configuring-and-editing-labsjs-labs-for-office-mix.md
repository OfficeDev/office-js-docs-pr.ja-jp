
# <a name="configuring-and-editing-labsjs-labs-for-office-mix"></a>Office Mix 用 LabsJS ラボの構成と編集



Office Mix は、ラボの構成の取得と設定のための office.js メソッドを提供しています。構成は、作成中のラボの種類と、ラボが返すデータ型を Office Mix に示します。この情報は、分析の収集と視覚化に使用されます。

## <a name="getting-the-lab-editor"></a>ラボ エディターを取得する

[Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) オブジェクトであるラボ エディターを使用すると、ラボを編集できるほか、ラボの構成の取得と設定ができます。ラボの編集が完了したら、 **Done** メソッドを呼び出す必要があります。ただし、編集中のラボを取り込んだり実行したりするとき以外は、 **Done** メソッドを呼び出す必要はありません。一度に開くことができるラボのインスタンスは 1 つのみであることに注意してください。

次のコードは、ラボ エディターの取得方法を示しています。




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

特定のラボ向けの構成を格納するには、 **Labs.LabEditor** で **getConfiguration** メソッドと [setConfiguration](../../../reference/office-mix/labs.labeditor.md) メソッドを使用します。構成 ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) は、どのデータがラボで収集および処理されるかを Office Mix に示します。構成には、名前、バージョン、その他の構成オプションなどの、ラボに関する一般的な情報が含まれています。構成の最も重要な部分は、ラボ コンポーネントの定義です。

次のコードは、構成の設定方法および取得方法を示しています。構成を設定するには、構成オブジェクトを作成してから、 **setConfiguration** メソッドを呼び出します。続いて、構成を取得するには、ラボ エディター オブジェクトに対して **getConfiguration** メソッドを呼び出します。




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## <a name="closing-the-editor"></a>エディターを閉じる

エディターを閉じるには、ラボの編集が完了した時点で、エディターで  **Done** メソッドを呼び出します。ラボの取り込みと編集の両方は行えないことに注意してください。しかし、 **Done** を呼び出した後で、ラボの編集または実行のいずれかを行うことができます。


## <a name="interacting-with-a-lab"></a>ラボを操作する

ラボの構成を設定すると、ラボの操作を開始できる状態になります。ラボを PowerPoint 内で実行する場合、操作がシミュレートされます。ラボを Office Mix レッスン プレーヤー内で実行する場合、データは Office Mix データベースに格納されて分析に使用されます。


### <a name="getting-the-lab-instance"></a>ラボのインスタンスを取得する

ラボの操作は、 [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md) オブジェクトを使用して行います。このオブジェクトは、現在のユーザー用に構成されたラボのインスタンスです。ラボの実行 (または「取り込み」) を行うには、 [Labs.takeLab](../../../reference/office-mix/labs.takelab.md) 関数を呼び出します。


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

インスタンスのオブジェクトには、コンポーネントのインスタンス ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md)、 [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) の配列が含まれています。これらは、構成で指定したコンポーネントにマップされます。実際、インスタンスとは、単なる変換されたバージョンの構成であり、サーバー側の ID をインスタンス オブジェクトにアタッチし、該当する場合に特定のフィールド (ヒントと答えなど) をユーザーに表示しないようにするために使用されます。


### <a name="managing-state"></a>状態を管理する

状態は、特定のラボを実行するユーザーに関連付けられた一時的な記憶域です。記憶域を使用すると、連続したラボの呼び出し間で情報を保持できます。たとえば、プログラミングのラボでは、ユーザーによる進行中の現在の作業を格納できます。

状態を  **set** するには、次のコードを使用します。




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

状態を  **get** するには、次のコードを使用します。




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## <a name="component-instances-and-results"></a>コンポーネントのインスタンスと結果

続いて、4 種類のコンポーネントのインスタンスを実装する方法について概要を説明します。また、コンポーネントのメソッドの簡単な例を示します。 

最初に、コンポーネントのインスタンスを操作する際は、2 つの主要概念をよく理解しておく必要があります。その 1 つが、 **試行** と **値** の概念です。

 **試行**

試行とは、ユーザーがコンポーネントのインスタンスを完了しようとすることです。たとえば、複数選択式の問題の場合、試行は、ユーザーが問題に取り組み始めた時に開始し、最終スコアが割り当てられた時に終了します。Office Mix の分析により、問題についてのユーザーの結果が集計されます。


 >**メモ**:  試行は、**DynamicComponent** 型を除くすべてのコンポーネント型で使用できます。

**getAttempts** メソッドを使用すると、特定のコンポーネントのインスタンスに関連付けられたすべての試行の結果を取得できます。結果を取得すると、ユーザーは、**resume** メソッドを使用して既存の試行のいずれかを再試行したり、**createAttempt** メソッドを使用して試行を新規作成したりできます。次の例は、そのプロセスを示しています。




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **値**

コンポーネントのインスタンスには、値の配列にマップされたキーの辞書が含まれています。この配列を使用して、コンポーネントに関連付けるヒント、フィードバック、またはその他すべての値のセットを格納できます。コンポーネントのインスタンスは、 **getValues** メソッドを使用してこれらの値へのアクセスを提供します。

たとえば、ヒントの値のクエリを実行すると、分析にユーザーがヒントを受け取ったというマークが付きます。値は試行ごとに追跡されます。

次のコード例は、ヒントのクエリの実行方法を示しています。




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### <a name="activitycomponentinstance"></a>ActivityComponentInstance


**ActivityComponentInstace** オブジェクトを使用して、ユーザーによるアクティビティ コンポーネントの操作を追跡します。このクラスは、ユーザーがアクティビティの操作を終了したことを示す **complete** メソッドを提供します。このメソッドは、ユーザーによる割り当てられたタスクの完了、読み取りの終了、またはアクティビティに関連付けられたその他すべてのエンド ポイントを示すことができます。次のコードは、**complete** メソッドの使用方法を示しています。


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### <a name="choicecomponentinstance"></a>ChoiceComponentInstance


**ChoiceComponentInstance** オブジェクトを使用して、ユーザーによる選択コンポーネントの操作を追跡します。選択コンポーネントとは、選択肢の一覧がユーザーに表示され、ユーザーがそこから選択をする形式の問題です。正答はある場合とない場合があります。このクラスには、主な 2 つのメソッド (**getSubmissions** と **submit**) があります。**getSubmissions** メソッドを使用すると、以前格納された送信内容を取得することができます。**submit** メソッドを使用すると、新しい送信内容を格納することができます。次のコード例は、これらのメソッドの使用方法を示しています。


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="inputcomponentinstance"></a>InputComponentInstance


**InputComponentInstance** オブジェクトを使用して、ユーザーによる入力コンポーネントの操作を追跡します。このクラスには、主な 2 つのメソッド (**getSubmission** と **submit**) があります。**getSubmissions** メソッドを使用すると、以前格納された送信内容を取得することができます。**submit** メソッドを使用すると、新しい送信内容を格納することができます。次のコード スニペットは、**getSubmissions** メソッドの使用方法を示しています。


```js
var submissions = this._attempt.getSubmissions();
```

**submit** メソッドを使用する際は、**InputComponentAnswer** オブジェクトが送信された答えを表していること、および **InputComponentResult** オブジェクトに結果が含まれていることに注意してください。戻り値は、答え、結果、および結果が送信された時刻を示すタイムスタンプを含む **InputComponentSubmission** オブジェクトです。




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="dynamiccomponentinstance"></a>DynamicComponentInstance


**DynamicComponentInstance** オブジェクトを使用して、ユーザーによる動的コンポーネントの操作を追跡します。このクラスの主なメソッドは、**getComponents**、**createComponent**、および **close** です。

次の例に示すように、 **getComponents** メソッドを使用すると、以前作成されたコンポーネントのインスタンスの一覧を取得できます。




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

次の例に示すように、 **createComponent** メソッドは、新しいコンポーネントを構築し、そのコンポーネントのインスタンスを返します。




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

**close** メソッドを使用して、コンポーネントを新規作成するための動的コンポーネントの使用が終了したことを示します。ブール型の **isClosed** メソッドを使用して、動的コンポーネントのインスタンスが閉じられたかどうかをテストすることも可能であることに注意してください。次のコード例は、**close** メソッドの使用方法を示しています。




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## <a name="additional-resources"></a>その他のリソース



- [Office Mix アドイン](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [チュートリアル:Office Mix 用の最初のラボを作成する](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
