// 全域變數
let questions = [];
let originalQuestions = [];
let currentQuestionIndex = 0;
let selectedOptions = {};
let quizStarted = false;
let quizSubmitted = false;
let timer;
let timeLeft;

// DOM 元素
const quizFileInput = document.getElementById('quiz-file');
const fileError = document.getElementById('file-error');
const fileSuccess = document.getElementById('file-success');
const startQuizBtn = document.getElementById('start-quiz');
const restartQuizBtn = document.getElementById('restart-quiz');
const continueQuizBtn = document.getElementById('continue-quiz');
const quizStartSection = document.getElementById('quiz-start');
const quizContainer = document.getElementById('quiz-container');
const questionDisplay = document.getElementById('question-display');
const prevQuestionBtn = document.getElementById('prev-question');
const nextQuestionBtn = document.getElementById('next-question');
const submitQuizBtn = document.getElementById('submit-quiz');
const earlySubmitQuizBtn = document.getElementById('early-submit-quiz');
const quizResultSection = document.getElementById('quiz-result');
const scoreDisplay = document.getElementById('score-display');
const resultMessage = document.getElementById('result-message');
const reviewQuizBtn = document.getElementById('review-quiz');
const exportResultsBtn = document.getElementById('export-results');
const quizSummary = document.getElementById('quiz-summary');
const quizInfo = document.getElementById('quiz-info');
const totalQuestionsDisplay = document.getElementById('total-questions');
const answeredDisplay = document.getElementById('answered');
const paletteContainer = document.getElementById('question-palette');
const timerDisplay = document.getElementById('timer-display');
const fileNameDisplay = document.getElementById('file-name');

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    checkForSavedQuiz();
    // 讓自訂按鈕能觸發隱藏的檔案上傳 input
    const customUploadBtn = document.querySelector('.custom-file-upload');
    if (customUploadBtn) {
        customUploadBtn.addEventListener('click', () => {
            quizFileInput.click();
        });
    }
});

// 事件監聽器
quizFileInput.addEventListener('change', handleFileUpload);
startQuizBtn.addEventListener('click', startQuiz);
restartQuizBtn.addEventListener('click', restartQuiz);
continueQuizBtn.addEventListener('click', continueQuiz);
prevQuestionBtn.addEventListener('click', showPreviousQuestion);
nextQuestionBtn.addEventListener('click', showNextQuestion);
submitQuizBtn.addEventListener('click', submitQuiz);
earlySubmitQuizBtn.addEventListener('click', () => {
    if (confirm('確定要提前交卷嗎？')) {
        submitQuiz();
    }
});
reviewQuizBtn.addEventListener('click', showSummary);
exportResultsBtn.addEventListener('click', exportResults);

// 函數
function handleFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        fileNameDisplay.textContent = '尚未選擇檔案';
        return;
    }
    
    fileNameDisplay.textContent = `已選擇: ${file.name}`;
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            if (jsonData.length === 0) {
                showError("Excel檔案中沒有找到題目");
                return;
            }
            
            originalQuestions = processQuizData(jsonData);
            questions = [...originalQuestions];

            startQuizBtn.disabled = false;
            fileSuccess.style.display = 'block';
            fileSuccess.textContent = '檔案已成功上傳！';
            fileError.style.display = 'none';
            quizInfo.style.display = 'block';
            totalQuestionsDisplay.textContent = questions.length;
            answeredDisplay.textContent = '0';
        } catch (error) {
            console.error(error);
            showError("讀取Excel檔案時發生錯誤，請檢查檔案格式");
        }
    };
    
    reader.onerror = function() {
        showError("讀取檔案時發生錯誤");
    };
    
    reader.readAsArrayBuffer(file);
}

function processQuizData(data) {
    if (data.length === 0) return [];
    const firstItem = data[0];

    // Case 1: Standard Multiple Choice (has "選項A", etc.)
    if (firstItem.hasOwnProperty('選項A') || firstItem.hasOwnProperty('選項B')) {
        const questionKey = firstItem.hasOwnProperty('問題') ? '問題' : '試題';
        const answerKey = firstItem.hasOwnProperty('正確答案') ? '正確答案' : '答案';

        return data.map((item, index) => {
            if (!item[questionKey] || item[questionKey].toString().trim() === '') return null;
            const question = {
                id: index + 1,
                text: item[questionKey],
                options: [],
                answer: (item[answerKey] || '').toString().trim().toUpperCase()
            };
            ['A', 'B', 'C', 'D', 'E'].forEach(key => {
                const optionValue = item[`選項${key}`];
                if (optionValue && optionValue.toString().trim() !== '') {
                    question.options.push({ key, text: optionValue.toString() });
                }
            });
            return question;
        }).filter(q => q !== null);
    }

    // Check for formats that use '試題' and '答案' columns
    if (firstItem.hasOwnProperty('試題') && firstItem.hasOwnProperty('答案')) {
        const firstAnswer = (firstItem['答案'] || '').toString().trim().toUpperCase();
        
        // Case 2: True/False (answer is O/X)
        if (['O', 'X', '○', '是', '非', 'TRUE', 'FALSE', 'V'].includes(firstAnswer)) {
            return data.map((item, index) => {
                 if (!item['試題'] || item['試題'].toString().trim() === '') return null;
                const question = {
                    id: index + 1,
                    text: item['試題'],
                    options: [ { key: 'O', text: 'O' }, { key: 'X', text: 'X' } ],
                    answer: ''
                };
                let rawAnswer = (item['答案'] || '').toString().trim().toUpperCase();
                if (['O', '○', 'TRUE', '是', 'V'].includes(rawAnswer)) {
                    question.answer = 'O';
                } else {
                    question.answer = 'X';
                }
                return question;
            }).filter(q => q !== null);
        } 
        
        // Case 3: Embedded Multiple Choice (answer is likely A/B/C/D...)
        else {
            return data.map((item, index) => {
                const rawQuestionText = item['試題'] || '';
                if (!rawQuestionText) return null;

                const options = [];
                const optionRegex = /\(([\w\d]+)\)(.*?)(?=\s*\([\w\d]+\)|$)/g;
                
                let questionText = rawQuestionText;
                let firstOptionIndex = rawQuestionText.search(/\s*\([\w\d]+\)/);

                if (firstOptionIndex !== -1) {
                    questionText = rawQuestionText.substring(0, firstOptionIndex).trim();
                }

                let match;
                while ((match = optionRegex.exec(rawQuestionText)) !== null) {
                    options.push({
                        key: match[1].toUpperCase(),
                        text: match[2].trim()
                    });
                }

                if (options.length === 0) return null;

                return {
                    id: index + 1,
                    text: questionText,
                    options: options,
                    answer: (item['答案'] || '').toString().trim().toUpperCase()
                };
            }).filter(q => q !== null);
        }
    }
    
    showError("無法識別的Excel格式。請檢查欄位名稱是否符合說明。");
    return [];
}

function showError(message) {
    fileError.textContent = message;
    fileError.style.display = 'block';
    fileSuccess.style.display = 'none';
    startQuizBtn.disabled = true;
}

function startQuiz() {
    clearSavedProgress();
    if (originalQuestions.length === 0) {
        showError("請先上傳有效的題庫檔案");
        return;
    }
    
    questions = [...originalQuestions];
    const randomCountInput = document.getElementById('random-count');
    const count = parseInt(randomCountInput.value, 10);

    if (count > 0 && count < questions.length) {
        // Fisher-Yates shuffle
        for (let i = questions.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [questions[i], questions[j]] = [questions[j], questions[i]];
        }
        questions = questions.slice(0, count);
    }

    // --- 計時器邏輯 ---
    const timeLimitInput = document.getElementById('time-limit');
    const timeLimit = parseInt(timeLimitInput.value, 10) * 60;
    if (timeLimit > 0) {
        timeLeft = timeLimit;
        timerDisplay.classList.remove('hidden');
        startTimer();
    } else {
        timerDisplay.classList.add('hidden');
    }
    // --------------------


    // 初始化答案
    selectedOptions = {};
    currentQuestionIndex = 0;
    quizSubmitted = false;
    
    quizStarted = true;
    quizStartSection.classList.add('hidden');
    quizContainer.classList.remove('hidden');
    quizInfo.classList.remove('hidden');
    restartQuizBtn.classList.remove('hidden');
    updateQuestionDisplay();
    updateProgressBar();
    updatePalette();
    
    document.getElementById('total-questions-display').textContent = questions.length;
}

function restartQuiz() {
    quizStarted = false;
    quizSubmitted = false;
    currentQuestionIndex = 0;
    selectedOptions = {};
    
    quizStartSection.classList.remove('hidden');
    quizContainer.classList.add('hidden');
    quizResultSection.classList.add('hidden');
    quizSummary.classList.add('hidden');
    quizInfo.classList.add('hidden');
    restartQuizBtn.classList.add('hidden');
    
    startQuizBtn.disabled = false;
    quizFileInput.value = '';
    fileSuccess.style.display = 'none';
    continueQuizBtn.classList.add('hidden');
    startQuizBtn.textContent = '開始新測驗';
    fileNameDisplay.textContent = '尚未選擇檔案';

    clearInterval(timer);
    timerDisplay.classList.add('hidden');

    if (paletteContainer) {
        paletteContainer.innerHTML = '';
    }
    clearSavedProgress();
}

function updateQuestionDisplay() {
    if (!quizStarted) return;
    
    const question = questions[currentQuestionIndex];
    
    // 更新問題編號顯示
    document.getElementById('current-question-number').textContent = currentQuestionIndex + 1;
    
    // 創建問題 HTML
    let optionsHtml = '';
    const isTrueFalse = question.options.length === 2 && question.options[0].key === 'O' && question.options[1].key === 'X';

    question.options.forEach(option => {
        const isSelected = selectedOptions[question.id] === option.key;
        let optionClass = 'option';
        if (quizSubmitted) {
            if (option.key === question.answer) {
                optionClass += ' correct';
            } else if (isSelected && option.key !== question.answer) {
                optionClass += ' incorrect';
            }
        } else if (isSelected) {
            optionClass += ' selected';
        }
        
        const displayText = isTrueFalse ? option.text : `${option.key}. ${option.text}`;

        optionsHtml += `
            <div class="${optionClass}" data-key="${option.key}" 
                 onclick="${!quizSubmitted ? `selectOption('${question.id}', '${option.key}')` : ''}">
                ${displayText}
            </div>
        `;
    });
    
    const questionHtml = `
        <div class="question-container">
            <div class="question">${question.id}. ${question.text}</div>
            <div class="options">${optionsHtml}</div>
        </div>
    `;
    
    questionDisplay.innerHTML = questionHtml;
    
    // 更新導航按鈕
    prevQuestionBtn.disabled = currentQuestionIndex === 0;
    earlySubmitQuizBtn.style.display = 'inline-block';
    
    if (currentQuestionIndex === questions.length - 1) {
        nextQuestionBtn.style.display = 'none';
        submitQuizBtn.style.display = 'inline-block';
        earlySubmitQuizBtn.style.display = 'none';
    } else {
        nextQuestionBtn.style.display = 'inline-block';
        submitQuizBtn.style.display = 'none';
    }
    
    // 更新已答題數
    const answeredCount = Object.keys(selectedOptions).length;
    answeredDisplay.textContent = answeredCount;
}

function selectOption(questionId, optionKey) {
    if (quizSubmitted) return;
    
    const questionIndex = questions.findIndex(q => q.id == questionId);
    if (questionIndex === -1) return;
    
    selectedOptions[questionId] = optionKey;
    updateQuestionDisplay();
    saveProgress();
    
    // 更新已答題數
    const answeredCount = Object.keys(selectedOptions).length;
    answeredDisplay.textContent = answeredCount;
    updatePalette();
}

function showPreviousQuestion() {
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        updateQuestionDisplay();
        updateProgressBar();
        updatePalette();
    }
}

function showNextQuestion() {
    if (currentQuestionIndex < questions.length - 1) {
        currentQuestionIndex++;
        updateQuestionDisplay();
        updateProgressBar();
        updatePalette();
    }
}

function updateProgressBar() {
    const progress = ((currentQuestionIndex + 1) / questions.length) * 100;
    document.getElementById('progress-bar').style.width = `${progress}%`;
}

function goToQuestion(index) {
    if (index >= 0 && index < questions.length) {
        currentQuestionIndex = index;
        updateQuestionDisplay();
        updateProgressBar();
        updatePalette();
    }
}

function updatePalette() {
    if (!quizStarted || !paletteContainer) return;

    paletteContainer.innerHTML = '';
    questions.forEach((question, index) => {
        const paletteItem = document.createElement('div');
        paletteItem.classList.add('palette-item');
        paletteItem.textContent = index + 1;
        
        if (selectedOptions[question.id]) {
            paletteItem.classList.add('answered');
        }

        if (index === currentQuestionIndex) {
            paletteItem.classList.add('current');
        }
        
        paletteItem.addEventListener('click', () => goToQuestion(index));
        paletteContainer.appendChild(paletteItem);
    });
}

function submitQuiz() {
    quizSubmitted = true;
    quizContainer.classList.add('hidden');
    quizResultSection.classList.remove('hidden');
    clearSavedProgress();
    clearInterval(timer);
    timerDisplay.classList.add('hidden');

    let correctAnswers = 0;
    questions.forEach(question => {
        if (selectedOptions[question.id] === question.answer) {
            correctAnswers++;
        }
    });

    const score = Math.round((correctAnswers / questions.length) * 100);
    scoreDisplay.textContent = `${score}%`;

    let message = '';
    if (score >= 90) {
        message = '太棒了！你真是個天才！';
    } else if (score >= 70) {
        message = '做得很好！繼續努力！';
    } else if (score >= 50) {
        message = '還有進步空間，加油！';
    } else {
        message = '別灰心，再試一次吧！';
    }
    resultMessage.textContent = message;

    prevQuestionBtn.disabled = true;
    nextQuestionBtn.disabled = true;
    submitQuizBtn.disabled = true;
}

function showSummary() {
    quizSummary.classList.toggle('hidden');
    if(quizSummary.classList.contains('hidden')) {
        reviewQuizBtn.textContent = "查看詳細結果";
    } else {
        reviewQuizBtn.textContent = "隱藏詳細結果";
    }


    let summaryHtml = '<h3>詳細報告</h3>';
    questions.forEach(question => {
        const selected = selectedOptions[question.id];
        const isCorrect = selected === question.answer;
        
        const correctAnswerOption = question.options.find(opt => opt.key === question.answer);
        const correctAnswerText = correctAnswerOption ? correctAnswerOption.text : '無有效答案';

        let selectedAnswerText = '未作答';
        if(selected) {
            const selectedAnswerOption = question.options.find(opt => opt.key === selected);
            selectedAnswerText = selectedAnswerOption ? selectedAnswerOption.text : '無效選項';
        }

        summaryHtml += `
            <div class="summary-item" style="border-left: 5px solid ${isCorrect ? '#2ecc71' : '#e74c3c'};">
                <div class="summary-question">${question.id}. ${question.text}</div>
                <p>你的答案：${selected || '未作答'} (${selectedAnswerText}) ${isCorrect ? '<span style="color: #2ecc71;">✔</span>' : '<span style="color: #e74c3c;">✖</span>'}</p>
                ${!isCorrect ? `<p>正確答案：${question.answer} (${correctAnswerText})</p>` : ''}
            </div>
        `;
    });

    quizSummary.innerHTML = summaryHtml;
} 

function exportResults() {
    let correctAnswers = 0;
    questions.forEach(question => {
        if (selectedOptions[question.id] === question.answer) {
            correctAnswers++;
        }
    });
    const score = Math.round((correctAnswers / questions.length) * 100);
    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    let reportHtml = `
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <title>測驗結果報告</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 20px auto; padding: 20px; border: 1px solid #ddd; }
            h1, h2 { color: #2c3e50; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; }
            .summary-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
            .summary-table th, .summary-table td { border: 1px solid #ddd; padding: 12px; text-align: left; }
            .summary-table th { background-color: #f5f7fa; }
            .score { font-size: 2em; color: #3498db; text-align: center; font-weight: bold; }
            .question-item { margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid #eee; }
            .question-text { font-weight: bold; margin-bottom: 10px; }
            .user-answer.incorrect { color: #e74c3c; font-weight: bold; }
            .correct-answer { color: #2ecc71; font-weight: bold; }
            footer { text-align: center; margin-top: 30px; font-size: 0.9em; color: #777; }
            @media print {
                body { border: none; box-shadow: none; }
                .no-print { display: none; }
            }
        </style>
    </head>
    <body>
        <h1>測驗結果報告</h1>
        <p>匯出時間: ${timestamp}</p>
        
        <h2>測驗總覽</h2>
        <table class="summary-table">
            <tr><th>總題數</th><td>${questions.length}</td></tr>
            <tr><th>答對題數</th><td>${correctAnswers}</td></tr>
            <tr><th>答錯題數</th><td>${questions.length - correctAnswers}</td></tr>
            <tr><th>得分</th><td class="score">${score}%</td></tr>
        </table>

        <h2>詳細報告</h2>
    `;

    questions.forEach((question, index) => {
        const selected = selectedOptions[question.id];
        const isCorrect = selected === question.answer;
        
        reportHtml += `
        <div class="question-item">
            <div class="question-text">${index + 1}. ${question.text}</div>
            <div>你的答案: <span class="user-answer ${!isCorrect && selected ? 'incorrect' : ''}">${selected || '未作答'}</span></div>
            <div>正確答案: <span class="correct-answer">${question.answer}</span></div>
        </div>
        `;
    });

    reportHtml += `
        <footer>報告由題庫測驗系統生成</footer>
        <script>
            // Optional: Add a print button
            const printBtn = document.createElement('button');
            printBtn.textContent = '列印此報告';
            printBtn.className = 'no-print';
            printBtn.style.cssText = 'display: block; margin: 20px auto; padding: 10px 20px; font-size: 1em; cursor: pointer;';
            printBtn.onclick = () => window.print();
            document.body.insertBefore(printBtn, document.querySelector('h2'));
        </script>
    </body>
    </html>`;

    const reportWindow = window.open('', '_blank');
    reportWindow.document.write(reportHtml);
    reportWindow.document.close();
}


// --- 計時器功能 ---

function startTimer() {
    timer = setInterval(() => {
        timeLeft--;
        updateTimerDisplay();
        if (timeLeft <= 0) {
            clearInterval(timer);
            alert('時間到！已自動為您交卷。');
            submitQuiz();
        }
    }, 1000);
}

function updateTimerDisplay() {
    if (!timerDisplay) return;
    const minutes = Math.floor(timeLeft / 60);
    let seconds = timeLeft % 60;
    seconds = seconds < 10 ? '0' + seconds : seconds;
    timerDisplay.textContent = `剩餘時間: ${minutes}:${seconds}`;
    
    if (timeLeft < 60) {
        timerDisplay.style.backgroundColor = '#e74c3c'; // 最後一分鐘變紅色
    } else {
        timerDisplay.style.backgroundColor = '#2c3e50';
    }
}


// --- 進度保存功能 ---

function saveProgress() {
    if (!quizStarted) return;
    const progress = {
        questions: questions,
        selectedOptions: selectedOptions,
        currentQuestionIndex: currentQuestionIndex,
        quizFileName: quizFileInput.files[0] ? quizFileInput.files[0].name : '已保存的測驗',
        timeLeft: timeLeft // 保存剩餘時間
    };
    localStorage.setItem('quizProgress', JSON.stringify(progress));
}

function checkForSavedQuiz() {
    const savedProgress = localStorage.getItem('quizProgress');
    if (savedProgress) {
        const progressData = JSON.parse(savedProgress);
        if (progressData && progressData.questions && progressData.questions.length > 0) {
            continueQuizBtn.classList.remove('hidden');
            startQuizBtn.textContent = '開始新測驗';
            fileSuccess.innerHTML = `偵測到上次未完成的測驗 <strong>(${progressData.quizFileName})</strong>。您可以繼續或開始新測驗。`;
            fileSuccess.style.display = 'block';
        }
    }
}

function continueQuiz() {
    const savedProgress = localStorage.getItem('quizProgress');
    if (!savedProgress) {
        showError("找不到儲存的進度。");
        return;
    }
    
    const progressData = JSON.parse(savedProgress);
    
    questions = progressData.questions;
    selectedOptions = progressData.selectedOptions;
    currentQuestionIndex = progressData.currentQuestionIndex;
    quizSubmitted = false;
    quizStarted = true;

    // 恢復計時器狀態
    if (progressData.timeLeft) {
        timeLeft = progressData.timeLeft;
        timerDisplay.classList.remove('hidden');
        startTimer();
    }

    quizStartSection.classList.add('hidden');
    quizContainer.classList.remove('hidden');
    quizInfo.classList.remove('hidden');
    restartQuizBtn.classList.remove('hidden');
    continueQuizBtn.classList.add('hidden');


    updateQuestionDisplay();
    updateProgressBar();
    updatePalette();
    
    document.getElementById('total-questions-display').textContent = questions.length;
    answeredDisplay.textContent = Object.keys(selectedOptions).length;
}

function clearSavedProgress() {
    localStorage.removeItem('quizProgress');
} 