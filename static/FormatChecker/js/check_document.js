document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("doc-check-form");
    const loading = document.getElementById("loading");
    const resultDiv = document.getElementById("result");
    const downloadBtn = document.getElementById("download-btn");

    const grammarCheck = document.getElementById("grammar_check");
    const exceptionSection = document.getElementById("exception-section");
    const exceptionInput = document.getElementById("exception_input");
    const addBtn = document.getElementById("add_exception");
    const exceptionWordsDiv = document.getElementById("exception-words");

    const fileInput = document.getElementById("document");
    const fileNameSpan = document.getElementById("file-name-display");

    const submitBtn = form.querySelector('button[type="submit"]');

    let exceptionWords = [];

    fileInput.addEventListener("change", () => {
        if (fileInput.files.length > 0) {
            fileNameSpan.textContent = fileInput.files[0].name;
        } else {
            fileNameSpan.textContent = "No file selected";
        }
    });

    grammarCheck.addEventListener("change", () => {
        exceptionSection.style.display = grammarCheck.checked ? "block" : "none";
    });

    exceptionInput.addEventListener("input", function () {
        exceptionInput.style.height = "auto";
        exceptionInput.style.height = exceptionInput.scrollHeight + "px";
    });

    addBtn.addEventListener("click", () => {
        const input = exceptionInput.value.trim();
        if (!input) return;

        const words = input.split(",").map(w => w.trim()).filter(w => w);

        let added = false;
        words.forEach(word => {
            if (word && !exceptionWords.includes(word)) {
                exceptionWords.push(word);
                added = true;
            }
        });

        if (added) {
            renderExceptions();
            exceptionInput.value = "";
        }
    });



    function renderExceptions() {
        exceptionWordsDiv.innerHTML = "";

        exceptionWords.forEach(word => {
            const span = document.createElement("span");
            span.className = "exception-item";
            span.textContent = word;

            const del = document.createElement("span");
            del.className = "delete-btn";
            del.textContent = "×";
            del.addEventListener("click", () => {
                exceptionWords = exceptionWords.filter(w => w !== word);
                renderExceptions();
            });

            span.appendChild(del);
            exceptionWordsDiv.appendChild(span);
        });
    }

    form.addEventListener("submit", async function(event) {
        event.preventDefault();

        submitBtn.disabled = true;
        submitBtn.textContent = "Checking...";

        loading.style.display = "block";
        resultDiv.innerText = "";
        downloadBtn.style.display = "none";

        let formData = new FormData(form);

        if (grammarCheck.checked && exceptionWords.length > 0) {
            formData.append("exception_words", JSON.stringify(exceptionWords));
        }

        try {
            let response = await fetch(form.action, {
                method: "POST",
                body: formData
            });

            let result = await response.json();
            console.log(result)
            loading.style.display = "none";

            submitBtn.disabled = false;
            submitBtn.textContent = "Check Document";

            if (result.error){
                resultDiv.innerHTML = "";
                const p = document.createElement("p");
                            console.log(result)
                            p.textContent = result.error;
                            resultDiv.appendChild(p);
            }
            else if(result && !result.error) {
                resultDiv.innerHTML = "";
                if (result.formatting) {
                    const formattingTitle = document.createElement("h2");
                    formattingTitle.textContent = "Formatting: ";
                    resultDiv.appendChild(formattingTitle);
                    result.formatting.forEach(line => {
                        if (line) {
                            const p = document.createElement("p");
                            p.textContent = line;
                            p.style.marginBottom = "5px";
                            resultDiv.appendChild(p);
                        }
                    })}
                if (result.grammar) {
                    const grammarTitle = document.createElement("h2");
                    grammarTitle.textContent = "Grammar: ";
                    resultDiv.appendChild(grammarTitle);
                    result.grammar.forEach(line => {
                        if (line) {
                            const p = document.createElement("p");
                            p.textContent = line;
                            p.style.marginBottom = "5px";
                            resultDiv.appendChild(p);
                        }
                    });
                }

                downloadBtn.style.display = "inline-block";
                downloadBtn.onclick = () => {
                    let content = "";

                    if (result.formatting && result.formatting.length > 0) {
                        content += "Formatting errors:\n" + result.formatting.join("\n") + "\n\n";
                    }

                    if (result.grammar && result.grammar.length > 0) {
                        content += "Grammar errors:\n" + result.grammar.join("\n") + "\n";
                    }

                    const blob = new Blob([content], { type: "text/plain" });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = "document_check_results.txt";
                    a.click();
                    URL.revokeObjectURL(url);
                };
            } else {
                resultDiv.innerText = "Something went wrong.";
            }
        } catch (error) {
            loading.style.display = "none";
            resultDiv.innerText = "Error: " + error.message;

            submitBtn.disabled = false;
            submitBtn.textContent = "Check Document";
        }
    });
});
