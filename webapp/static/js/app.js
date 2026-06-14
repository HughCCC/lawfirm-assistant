/**
 * Legal Report Generator — Frontend Logic
 *
 * Handles: provider selection, form state, API calls,
 * loading/results display, and download triggering.
 */

document.addEventListener("DOMContentLoaded", () => {
    // ─── DOM References ──────────────────────────────────────────

    const formState = document.getElementById("form-state");
    const resultState = document.getElementById("result-state");
    const progressPanel = document.getElementById("progress-panel");

    const providerRadios = document.querySelectorAll('input[name="provider"]');
    const modelSelect = document.getElementById("model-select");
    const baseUrlGroup = document.getElementById("base-url-group");
    const baseUrlInput = document.getElementById("base-url");
    const apiKeyInput = document.getElementById("api-key");
    const questionTextarea = document.getElementById("question");
    const charCount = document.getElementById("char-count");

    const btnGenerate = document.getElementById("btn-generate");
    const btnClear = document.getElementById("btn-clear");
    const btnText = btnGenerate.querySelector(".btn-text");
    const btnSpinner = btnGenerate.querySelector(".btn-spinner");

    // Result elements
    const resultSuccess = document.getElementById("result-success");
    const resultFailure = document.getElementById("result-failure");
    const resultTitle = document.getElementById("result-title");
    const resultRetries = document.getElementById("result-retries");
    const btnDownload = document.getElementById("btn-download");
    const errorMessage = document.getElementById("error-message");
    const btnNewReport = document.getElementById("btn-new-report");
    const btnRetry = document.getElementById("btn-retry");
    const btnBackForm = document.getElementById("btn-back-form");

    // Progress ring
    const progressPercent = document.getElementById("progress-percent");
    const progressRingFill = document.querySelector(".progress-ring-fill");

    // ─── Progress Simulation State ─────────────────────────────

    const RING_CIRCUMFERENCE = 326.73;  // 2 * PI * 52
    let progressTimer = null;
    let progressStartTime = null;
    const TARGET_DURATION_MS = 90000;   // 90 seconds target

    // ─── Provider Configurations ────────────────────────────────

    const PROVIDER_CONFIG = {
        openai: {
            models: ["gpt-4o", "gpt-4-turbo", "gpt-4o-mini", "o3-mini"],
            defaultModel: "gpt-4o",
            hint: "",
        },
        anthropic: {
            models: ["claude-sonnet-4-20250514", "claude-opus-4-20250514", "claude-haiku-4-5-20251001"],
            defaultModel: "claude-sonnet-4-20250514",
            hint: "推荐使用 Claude Sonnet 4 或 Opus 4 以获得最佳法律分析质量",
        },
        deepseek: {
            models: ["deepseek-chat", "deepseek-reasoner"],
            defaultModel: "deepseek-chat",
            hint: "推荐使用 deepseek-chat（中文法律分析质量优秀）",
        },
        custom: {
            models: [],
            defaultModel: "",
            hint: "请输入与 OpenAI API 兼容的 Base URL",
        },
    };

    // ─── Provider Selection ──────────────────────────────────────

    function getSelectedProvider() {
        const checked = document.querySelector('input[name="provider"]:checked');
        return checked ? checked.value : "openai";
    }

    function updateProviderUI() {
        const provider = getSelectedProvider();
        const config = PROVIDER_CONFIG[provider];

        // Model select
        modelSelect.innerHTML = '<option value="">使用默认模型</option>';
        config.models.forEach((m) => {
            const option = document.createElement("option");
            option.value = m;
            option.textContent = m;
            if (m === config.defaultModel) option.selected = true;
            modelSelect.appendChild(option);
        });

        // Base URL visibility
        baseUrlGroup.style.display = provider === "custom" ? "" : "none";
        if (provider !== "custom") baseUrlInput.value = "";

        // Model select visibility
        const hasModels = config.models.length > 0;
        modelSelect.closest(".form-group").style.display = hasModels ? "" : "block";
    }

    providerRadios.forEach((radio) => {
        radio.addEventListener("change", updateProviderUI);
    });

    // Initialize
    updateProviderUI();

    // ─── Character Count ─────────────────────────────────────────

    questionTextarea.addEventListener("input", () => {
        const count = questionTextarea.value.length;
        charCount.textContent = `${count} 字符`;
        charCount.style.color = count < 20 ? "var(--color-error)" : "";
    });

    // ─── Button: Clear ──────────────────────────────────────────

    btnClear.addEventListener("click", () => {
        apiKeyInput.value = "";
        questionTextarea.value = "";
        charCount.textContent = "0 字符";
        charCount.style.color = "";
        document.querySelector('input[value="openai"]').checked = true;
        updateProviderUI();
    });

    // ─── Progress Simulation ──────────────────────────────────────

    /**
     * Ease-out cubic: fast start, slow finish, maxes out at 90%.
     * Formula: progress(t) = 0.9 * (1 - (1-t)^3)
     */
    function easeOutProgress(elapsed, total) {
        const t = Math.min(elapsed / total, 1.0);
        return 0.9 * (1 - Math.pow(1 - t, 3));
    }

    function startProgress() {
        progressStartTime = Date.now();
        progressRingFill.setAttribute("stroke-dashoffset", RING_CIRCUMFERENCE);
        progressRingFill.classList.remove("completed");
        progressPercent.classList.remove("completed");
        updateProgressDisplay(0);

        progressTimer = setInterval(() => {
            const elapsed = Date.now() - progressStartTime;
            const progress = easeOutProgress(elapsed, TARGET_DURATION_MS);
            updateProgressDisplay(progress);
        }, 200);
    }

    function updateProgressDisplay(progress) {
        const percent = Math.round(progress * 100);
        progressPercent.textContent = percent + "%";
        const offset = RING_CIRCUMFERENCE * (1 - progress);
        progressRingFill.setAttribute("stroke-dashoffset", offset);
    }

    function completeProgress() {
        if (progressTimer) {
            clearInterval(progressTimer);
            progressTimer = null;
        }
        updateProgressDisplay(1.0);
        progressRingFill.classList.add("completed");
        progressPercent.classList.add("completed");
    }

    function stopProgress() {
        if (progressTimer) {
            clearInterval(progressTimer);
            progressTimer = null;
        }
    }

    // ─── Button: Generate ───────────────────────────────────────

    btnGenerate.addEventListener("click", async () => {
        // Validate
        const question = questionTextarea.value.trim();
        if (question.length < 20) {
            alert("请输入至少 20 字符的法律问题描述。");
            questionTextarea.focus();
            return;
        }

        const apiKey = apiKeyInput.value.trim();
        if (apiKey.length < 10) {
            alert("请输入有效的 API Key。");
            apiKeyInput.focus();
            return;
        }

        const provider = getSelectedProvider();

        // Validate base URL for custom provider
        if (provider === "custom") {
            const baseUrl = baseUrlInput.value.trim();
            if (!baseUrl) {
                alert("自定义提供商需要填写 Base URL。");
                baseUrlInput.focus();
                return;
            }
        }

        // Enter loading state
        setLoadingState(true);
        hideResults();

        const model = modelSelect.value || null;
        const baseUrl = provider === "custom" ? baseUrlInput.value.trim() : null;

        const payload = {
            question,
            provider,
            api_key: apiKey,
            model,
            base_url: baseUrl,
        };

        try {
            const response = await fetch("/api/generate", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });

            const data = await response.json();

            // Snap progress to 100% briefly before showing result
            completeProgress();
            await new Promise(resolve => setTimeout(resolve, 600));

            if (data.success) {
                showSuccess(data);
            } else {
                showFailure(data);
            }
        } catch (err) {
            stopProgress();
            showFailure({
                success: false,
                error: `网络请求失败：${err.message}。请检查服务器是否在运行（http://localhost:8000）。`,
                retries_used: 0,
            });
        } finally {
            setLoadingState(false);
        }
    });

    // ─── Loading State ───────────────────────────────────────────

    function setLoadingState(loading) {
        if (loading) {
            btnText.style.display = "none";
            btnSpinner.style.display = "inline-flex";
            btnGenerate.disabled = true;
            btnClear.disabled = true;
            providerRadios.forEach((r) => (r.disabled = true));
            modelSelect.disabled = true;
            apiKeyInput.disabled = true;
            questionTextarea.disabled = true;
            progressPanel.style.display = "";
            resultState.style.display = "none";
            startProgress();
        } else {
            btnText.style.display = "";
            btnSpinner.style.display = "none";
            btnGenerate.disabled = false;
            btnClear.disabled = false;
            providerRadios.forEach((r) => (r.disabled = false));
            modelSelect.disabled = false;
            apiKeyInput.disabled = false;
            questionTextarea.disabled = false;
            progressPanel.style.display = "none";
            stopProgress();
        }
    }

    // ─── Results Display ─────────────────────────────────────────

    function hideResults() {
        resultState.style.display = "none";
        resultSuccess.style.display = "none";
        resultFailure.style.display = "none";
        formState.style.display = "";
    }

    function showSuccess(data) {
        formState.style.display = "none";
        resultState.style.display = "";
        resultSuccess.style.display = "";
        resultFailure.style.display = "none";

        resultTitle.textContent = `报告标题：${data.title}`;
        resultRetries.textContent =
            data.retries_used > 0
                ? `（经过 ${data.retries_used} 次修正后生成）`
                : "";

        btnDownload.href = data.docx_url;
        btnDownload.download = data.docx_filename;
    }

    function showFailure(data) {
        formState.style.display = "none";
        resultState.style.display = "";
        resultSuccess.style.display = "none";
        resultFailure.style.display = "";

        errorMessage.textContent = data.error || "未知错误，请重试。";
    }

    // ─── Result Actions ──────────────────────────────────────────

    btnNewReport.addEventListener("click", () => {
        hideResults();
        formState.style.display = "";
    });

    btnRetry.addEventListener("click", () => {
        hideResults();
        formState.style.display = "";
        // Re-trigger generation
        btnGenerate.click();
    });

    btnBackForm.addEventListener("click", () => {
        hideResults();
        formState.style.display = "";
    });

    // ─── Keyboard Shortcut: Cmd/Ctrl+Enter to generate ──────────

    questionTextarea.addEventListener("keydown", (e) => {
        if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
            e.preventDefault();
            btnGenerate.click();
        }
    });

    // ─── Format Upload ──────────────────────────────────────────

    const fileInput = document.getElementById("format-file");
    const btnFormat = document.getElementById("btn-format");
    const formatResult = document.getElementById("format-result");
    const formatStatus = document.getElementById("format-status");
    const formatDownload = document.getElementById("format-download");

    if (!btnFormat || !fileInput) {
        console.error("Format elements not found: btnFormat=" + !!btnFormat + " fileInput=" + !!fileInput);
    }

    btnFormat.addEventListener("click", async () => {
        console.log("Format button clicked");
        const file = fileInput.files[0];
        if (!file) {
            formatStatus.textContent = "⚠ 请先选择一个 .docx 文件";
            formatStatus.style.color = "var(--color-error)";
            formatResult.style.display = "flex";
            formatDownload.style.display = "none";
            return;
        }

        if (!file.name.endsWith(".docx")) {
            formatStatus.textContent = "⚠ 仅支持 .docx 格式的 Word 文档";
            formatStatus.style.color = "var(--color-error)";
            formatResult.style.display = "flex";
            formatDownload.style.display = "none";
            return;
        }

        // Enter loading state
        btnFormat.disabled = true;
        btnFormat.textContent = "正在调整格式…";
        formatResult.style.display = "flex";
        formatStatus.textContent = "⏳ 正在上传并处理…";
        formatStatus.style.color = "var(--color-navy)";
        formatDownload.style.display = "none";

        const formData = new FormData();
        formData.append("file", file);

        try {
            const response = await fetch("/api/format", {
                method: "POST",
                body: formData,
            });

            const data = await response.json();
            console.log("Format response:", data);

            if (data.success) {
                formatStatus.textContent = "✓ 格式调整完成";
                formatStatus.style.color = "var(--color-success)";
                formatDownload.href = data.url;
                formatDownload.download = data.filename;
                formatDownload.style.display = "";
            } else {
                formatStatus.textContent = "✗ 格式调整失败：" + (data.detail || "未知错误");
                formatStatus.style.color = "var(--color-error)";
                formatDownload.style.display = "none";
            }
        } catch (err) {
            console.error("Format error:", err);
            formatStatus.textContent = "✗ 网络请求失败：" + err.message;
            formatStatus.style.color = "var(--color-error)";
            formatDownload.style.display = "none";
        } finally {
            btnFormat.disabled = false;
            btnFormat.textContent = "上传并调整格式";
        }
    });

    // Reset format UI on new file selection
    fileInput.addEventListener("change", () => {
        formatResult.style.display = "none";
        formatStatus.style.color = "";
        formatDownload.style.display = "";
    });
});
