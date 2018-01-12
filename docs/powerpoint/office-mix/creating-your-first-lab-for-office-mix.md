
# <a name="walkthrough-creating-your-first-lab-for-office-mix"></a>チュートリアル: Office Mix 用の最初のラボの作成
ステップ バイ ステップ チュートリアルを使用して、最初の LabsJS ラボを作成します。



このチュートリアルでは、簡単な LabsJS ラボを一から作成します。ラボは、1 つの質問のみ提供する単純な ○× クイズとなります。 

Visual Studio プロジェクト テンプレートで作業を開始するのではなく、3 つの空のファイルで作業を開始します。これは、ラボがどれだけ簡単であるかを示しています。 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
Visual Studio のテンプレートで開始していないため、これらのファイルを編集する際には、任意のコード エディターを使用できます。実際、HTML ファイルは簡易であり、お好みで、チュートリアル ファイルから HTML マークアップをコピーして、貼り付けることができます。ただし、HTML5 を使用する必要があるため、DOCTYPE 宣言が `<!DOCTYPE html>` であることを確認してください。CSS ファイルは、オプションです。すべての面倒な作業は、JavaScript (.js) ファイル、TrueFalse.js で行われます。このチュートリアルでは、次の 4 つの主なラボの機能をカバーします。

- セットアップ (ホストへの接続)
    
- モードの変更 (編集モードと表示モード)
    
- ラボの編集
    
- ラボの実行
    

 **メモ**  
 ---
 ファイル labhost.html は、Web サーバー上で実行され、ラボの開発とテスト用のホスティング環境を提供します。これにより、ラボの開発が大幅に簡略化されます。開発環境のセットアップについては、「[Office Mix 用の LabsJS を開始する](get-started-with-labsjs-for-office-mix.md)」をご覧ください。<br/><br/>

最後に、この SDK で配布されたファイルの中に、完成した JavaScript ファイル (TrueFalse.js) が表示されます。この後に、コーディング プロセスのチュートリアルが続きます。

## <a name="connecting-to-the-lab-host"></a>ラボのホストに接続する

この環境のラボは、(開発およびテスト用の) ラボ ホスト、または Office.js ホストによって提供される既定のランタイム ホストのいずれかを使用して実行できます。opening 関数は、単純な if/else 式を使用して、これらのホスティング コンテキストのうちどれが適用されているかをテストします。


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

**PostMessageLabHost** オブジェクトは、labhost.html の開発環境で実行します。一方、運用環境では、ラボは **OfficeJSLabHost** を使用して PowerPoint/Office Mix で実行します。

次に、渡された jQuery の遅延オブジェクトの解決または拒否を行うコールバックを作成するヘルパー メソッドを作成します。メソッド **createCallback** を使用して、jQuery の promises から labs.js が定義したコールバックに移動します。




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

また、特定の質問と答えに関するラボ構成を取得するヘルパー メソッドも作成します。




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## <a name="mode-changes"></a>モードの変更

ラボは、常に **view** および **edit** という 2 つの状態 (モード) のいずれかです。そのため、クイズのために状態と動作をキャプチャおよび保持する方法が必要になります。このためにクラスを作成します。


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

さらに、クイズの質問に対する答え (つまり、"submission") が正しいかどうかに基づいてクイズの UI を更新するヘルパー メソッドを用意します。




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

さらに、編集モードと表示モードを切り替える関数が必要になります。




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

次の関数は、UI から受け取った変更イベントに基づいてクイズの構成を更新します。




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

次に、特定の質問と答えに基づいて、サーバーに格納されたラボ構成データを更新します。




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

次に、編集モードのラボで行われた更新内容を、これまでに行った構成変更にバインドする関数があります。その後に、以前バインドした変更ハンドラーのバインドを解除するコードが続きます。




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

ここからが、セクションの主要部分、つまり表示モードと編集モードを相互に切り替えるメソッドです。表示モードから編集モードへの切り替えから始めます。




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

ここからは、編集モードから表示モードへの切り替えです。




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

最後に、ホストに接続してドキュメントの準備ができたら、クイズを開始します。




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## <a name="additional-resources"></a>その他のリソース
<a name="bk_addresources"> </a>


- [Office Mix アドイン](office-mix-add-ins.md)
    
