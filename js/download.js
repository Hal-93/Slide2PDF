function waitForScriptsToLoad(callback) {
    if (
        typeof html2pdf !== "undefined" &&
        typeof html2canvas !== "undefined" &&
        typeof window.jspdf !== "undefined"
    ) {
        callback();
    } else {
        setTimeout(() => waitForScriptsToLoad(callback), 100);
    }
}

waitForScriptsToLoad(main);

function main() {
    let isBlockingUserActions = false;
    let userActionHandlersAttached = false;

    function htmlspecialchars(unsafeText) {
        if (typeof unsafeText !== "string") {
            return unsafeText;
        }
        return unsafeText.replace(/[&'`"<>]/g, function (match) {
            return {
                "&": "&amp;",
                "'": "&#x27;",
                "`": "&#x60;",
                '"': "&quot;",
                "<": "&lt;",
                ">": "&gt;",
            }[match];
        });
    }

    async function captureSlide(slide) {
        return new Promise((resolve, reject) => {
            console.log("Capturing slide", slide);

            html2canvas(slide)
                .then((canvas) => {
                    console.log("Captured canvas size:", canvas.width, canvas.height);
                    resolve({
                        dataUrl: canvas.toDataURL("image/jpeg", 0.92),
                        width: canvas.width,
                        height: canvas.height,
                    });
                })
                .catch((err) => {
                    console.error("html2canvas error", err);
                    reject(err);
                });
        });
    }

    function downloadPDF(slides, filename) {
        const { jsPDF } = window.jspdf;

        if (!slides || slides.length === 0) {
            console.error("No slides to export.");
            return;
        }

        const first = slides[0];

        const pxToPt = (px) => (px * 72) / 96;

        const firstWidthPt = pxToPt(first.width);
        const firstHeightPt = pxToPt(first.height);

        const orientation =
            first.width >= first.height ? "landscape" : "portrait";

        const pdf = new jsPDF({
            orientation,
            unit: "pt",
            format: [firstWidthPt, firstHeightPt],
        });

        slides.forEach((slide, index) => {
            const { dataUrl, width, height } = slide;
            const wPt = pxToPt(width);
            const hPt = pxToPt(height);

            if (index > 0) {
                pdf.addPage([wPt, hPt]);
            }

            pdf.addImage(dataUrl, "JPEG", 0, 0, wPt, hPt);
        });

        pdf.save(filename);
    }

    function showLoadingOverlay(progress) {
        const overlay = document.createElement("div");
        overlay.id = "loading-overlay";
        overlay.style.position = "fixed";
        overlay.style.top = "0";
        overlay.style.left = "0";
        overlay.style.width = "100%";
        overlay.style.height = "100%";
        overlay.style.backgroundColor = "rgba(0, 0, 0, 0.9)";
        overlay.style.zIndex = "10000";
        overlay.style.pointerEvents = "all";
        overlay.style.color = "white";
        overlay.style.display = "flex";
        overlay.style.justifyContent = "center";
        overlay.style.alignItems = "center";
        overlay.style.fontSize = "24px";
        overlay.innerHTML = `<div>ダウンロード中: ${progress}%</div>`;
        document.body.appendChild(overlay);
    }

    function updateLoadingOverlay(progress) {
        const overlay = document.getElementById("loading-overlay");
        if (overlay) {
            overlay.innerHTML = `<div>ダウンロード中: ${progress}%</div>`;
        }
    }

    function hideLoadingOverlay() {
        const overlay = document.getElementById("loading-overlay");
        if (overlay) {
            document.body.removeChild(overlay);
        }
        isBlockingUserActions = false;
        detachUserActionBlockers();
    }

    function init() {
        if (!window.location.href.match(/docs\.google\.com\/presentation\/d\/(e|)/)) {
            return;
        }

        if (document.getElementById("slide2pdf-download-btn")) {
            return;
        }

        const pdfBtn = createDownloadButton();
        attachButtonToUI(pdfBtn);

        pdfBtn.onclick = () => {
            const name_html = document.title.replace(/\s*\[.*?\]\s*/g, "") + ".pdf";

            isBlockingUserActions = true;
            attachUserActionBlockers();
            showLoadingOverlay(0);
            processSlides(name_html);
        };

        async function processSlides(filename) {
            console.log("processSlides started");

            const startingSlide = getCurrentSlideNumber();
            let totalSlides = getTotalSlides();
            if (!totalSlides || Number.isNaN(totalSlides)) {
                console.error("スライドが検出できませんでした。");
                hideLoadingOverlay();
                return;
            }

            let currentSlide = getCurrentSlideNumber();
            let slides_data = [];

            if (currentSlide !== 1) {
                await navigateToSlide(1);
                currentSlide = 1;
            }

            try {
                for (let i = 0; i < totalSlides; i++) {
                    let slide = getActiveSlideElement();
                    if (!slide) break;

                    try {
                        let image = await captureSlide(slide);
                        slides_data.push(image);
                        updateLoadingOverlay(
                            Math.round((slides_data.length / totalSlides) * 100)
                        );
                    } catch (error) {
                        console.error("エラー: ", error);
                    }

                    if (slides_data.length >= totalSlides) break;

                    if (i < totalSlides - 1) {
                        const target = currentSlide + 1;
                        await navigateToSlide(target);
                        currentSlide = target;
                    }
                }

                downloadPDF(slides_data, filename);
            } finally {
                await navigateToSlide(startingSlide || 1);
                hideLoadingOverlay();
            }
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init);
    } else {
        init();
    }

    function getCurrentSlideNumber() {
        const url = window.location.href;
        const match = url.match(/p(\d+)$/);
        if (match) {
            return parseInt(match[1], 10);
        }

        const active = getActiveSlideElement();
        const ariaPos = active
            ? parseInt(active.getAttribute("aria-posinset") || "", 10)
            : null;

        return ariaPos || 1;
    }

    function goToSlide(slideNumber) {
        window.location.hash = "slide=id.p" + slideNumber;
    }

    function sleep(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    async function navigateToSlide(slideNumber) {
        const previousNumber = getCurrentSlideNumber();
        goToSlide(slideNumber);

        const maxAttempts = 50;
        for (let i = 0; i < maxAttempts; i++) {
            const currentFromHash = getCurrentSlideNumber();
            const active = getActiveSlideElement();
            const currentFromAria = active
                ? parseInt(active.getAttribute("aria-posinset") || "", 10)
                : null;

            if (
                currentFromHash === slideNumber ||
                currentFromAria === slideNumber
            ) {
                return;
            }

            if (i % 10 === 0 && currentFromHash === previousNumber) {
                dispatchArrowNavigation(slideNumber > previousNumber ? "ArrowRight" : "ArrowLeft");
            }

            await sleep(150);
        }

        console.warn("Slide navigation timed out");
    }

    function dispatchArrowNavigation(key) {
        const eventInit = {
            key,
            code: key,
            keyCode: key === "ArrowRight" ? 39 : 37,
            which: key === "ArrowRight" ? 39 : 37,
            bubbles: true,
        };

        document.dispatchEvent(new KeyboardEvent("keydown", eventInit));
        document.dispatchEvent(new KeyboardEvent("keyup", eventInit));
    }

    function getTotalSlides() {
        const caption = document.querySelector(
            ".docs-material-menu-button-flat-default-caption"
        );
        const size = caption ? caption.getAttribute("aria-setsize") : null;

        if (size) {
            return parseInt(size, 10);
        }

        const svgSlides = document.querySelectorAll(
            ".punch-viewer-svgpage-svgcontainer"
        );

        return svgSlides.length;
    }

    function getActiveSlideElement() {
        const svgSlides = document.querySelectorAll(
            ".punch-viewer-svgpage-svgcontainer"
        );

        return svgSlides[svgSlides.length - 1] || null;
    }

    function createDownloadButton() {
        const pdfBtn = document.createElement("button");
        pdfBtn.id = "slide2pdf-download-btn";
        pdfBtn.setAttribute("aria-label", "Slide2PDF download");
        pdfBtn.style.width = "32px";
        pdfBtn.style.height = "32px";
        pdfBtn.style.borderRadius = "50%";
        pdfBtn.style.border = "none";
        pdfBtn.style.padding = "0";
        pdfBtn.style.backgroundColor = "#1a73e8";
        pdfBtn.style.display = "inline-flex";
        pdfBtn.style.alignItems = "center";
        pdfBtn.style.justifyContent = "center";
        pdfBtn.style.cursor = "pointer";
        pdfBtn.style.boxShadow = "0 2px 6px rgba(0, 0, 0, 0.18)";
        pdfBtn.style.transition = "transform 0.15s ease, box-shadow 0.15s ease";
        pdfBtn.onmouseenter = () => {
            pdfBtn.style.transform = "translateY(-1px)";
            pdfBtn.style.boxShadow = "0 4px 10px rgba(0, 0, 0, 0.26)";
        };
        pdfBtn.onmouseleave = () => {
            pdfBtn.style.transform = "translateY(0)";
            pdfBtn.style.boxShadow = "0 2px 6px rgba(0, 0, 0, 0.18)";
        };
        pdfBtn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M12 4v10.5m0 0l-4-4m4 4l4-4" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M6 18h12" stroke="white" stroke-width="2" stroke-linecap="round"/>
            </svg>
        `;
        return pdfBtn;
    }

    function styleForControlBar(btn) {
        btn.style.position = "static";
        btn.style.marginLeft = "6px";
        btn.style.flex = "0 0 auto";
        btn.style.verticalAlign = "middle";
        btn.style.alignSelf = "center";
        btn.style.lineHeight = "1";
    }

    function styleAsFloatingButton(btn) {
        btn.style.position = "fixed";
        btn.style.top = "20px";
        btn.style.right = "20px";
        btn.style.zIndex = "99999";
    }

    function findControlBarContainer() {
        const selectors = [
            ".punch-viewer-controls",
            ".punch-viewer-navbar",
            ".punch-viewer-navbar-lower",
            ".punch-viewer-action-bar",
            ".punch-viewer-controls-panel",
            "div[role='toolbar']",
        ];

        for (const sel of selectors) {
            const node = document.querySelector(sel);
            if (node) return node;
        }

        const optBtn = findOptionsButton(document);
        if (optBtn && optBtn.parentElement) {
            return optBtn.parentElement;
        }

        return null;
    }

    function findOptionsButton(root) {
        const labels = ["options", "オプション", "その他の操作"];
        const candidates = root.querySelectorAll("button, div[role='button']");
        for (const btn of candidates) {
            const aria = (btn.getAttribute("aria-label") || "").toLowerCase();
            const tooltip = (btn.getAttribute("data-tooltip") || "").toLowerCase();
            const title = (btn.getAttribute("title") || "").toLowerCase();
            const className = btn.className || "";

            const matchesLabel =
                labels.some((label) => aria.includes(label)) ||
                labels.some((label) => tooltip.includes(label)) ||
                labels.some((label) => title.includes(label));

            const matchesClass = className.includes("punch-viewer-navbar-more");

            if (matchesLabel || matchesClass) {
                return btn;
            }
        }
        return null;
    }

    function attachButtonToUI(btn) {
        let attempts = 0;
        const maxAttempts = 15;

        const tryAttach = () => {
            if (btn.isConnected) return;

            const controlBar = findControlBarContainer();
            if (controlBar) {
                styleForControlBar(btn);
                const optionsButton = findOptionsButton(controlBar);
                if (optionsButton && optionsButton.parentElement) {
                    optionsButton.insertAdjacentElement("afterend", btn);
                } else {
                    controlBar.appendChild(btn);
                }
                return;
            }

            attempts += 1;
            if (attempts < maxAttempts) {
                setTimeout(tryAttach, 200);
            } else {
                styleAsFloatingButton(btn);
                if (!btn.isConnected) {
                    document.body.appendChild(btn);
                }
            }
        };

        tryAttach();
    }

    function attachUserActionBlockers() {
        if (userActionHandlersAttached) return;

        const prevent = (e) => {
            if (!isBlockingUserActions) return;
            e.stopPropagation();
            e.preventDefault();
        };

        const types = [
            "keydown",
            "keypress",
            "keyup",
            "mousedown",
            "mouseup",
            "click",
            "dblclick",
            "contextmenu",
            "wheel",
            "touchstart",
            "touchend",
            "pointerdown",
            "pointerup",
        ];

        types.forEach((type) => {
            document.addEventListener(type, prevent, true);
        });

        document._slide2pdfPreventFn = prevent;
        document._slide2pdfPreventTypes = types;
        userActionHandlersAttached = true;
    }

    function detachUserActionBlockers() {
        if (!userActionHandlersAttached) return;
        const prevent = document._slide2pdfPreventFn;
        const types = document._slide2pdfPreventTypes;
        if (prevent && types) {
            types.forEach((type) => {
                document.removeEventListener(type, prevent, true);
            });
        }
        delete document._slide2pdfPreventFn;
        delete document._slide2pdfPreventTypes;
        userActionHandlersAttached = false;
    }
}
