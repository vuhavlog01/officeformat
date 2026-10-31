/**
 * PHẦN MỀM LẬP & XUẤT LỊCH CÔNG TÁC TUẦN
 * Ban Thường vụ Đảng ủy / UBND
 * Frontend Logic & Real-time Live Preview Engine
 */

// State toàn cục của ứng dụng
let appState = {
    selectedDate: new Date(), // Ngày mốc
    currentWeekOffset: 1,     // Mặc định: 1 = Tuần kế tiếp
    zoomLevel: 1.0,
    includeSignature: true,
    autoMarkRedOnCurrentWeek: true, // Tự động tô đỏ lịch mới khi ở tuần hiện tại
    baseSnapshot: null,             // Bản gốc để so sánh thay đổi
    data: {
        agency_upper: "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG",
        agency_name: "ĐẢNG UỶ XÃ QUẾ SƠN",
        doc_date_text: "Quế Sơn, ngày 01 tháng 11 năm 2026",
        doc_title: "lịch công tác của Ban Thường vụ Đảng ủy từ ngày 26/10/2026 đến ngày 01/11/2026",
        days: [],
        driver_duty_note: "* Thứ Bảy, Chủ nhật, ngày 31/10 – 01/11/2026. A.Lực trực lái xe Đảng ủy.",
        invitation_note: "* Lịch công tác thay cho giấy mời họp Thường trực, Ban Thường vụ Đảng ủy.",
        change_note: "", // Chú thích lịch thay đổi nếu có
        next_week_title: "* Tuần 45:",
        next_week_items: [],
        recipients: [
            "Ủy viên BTV Đảng ủy;",
            "Thường trực HĐND, UBND xã;",
            "Các cơ quan CTTM, GV Đảng ủy;",
            "Lưu Văn phòng Đảng ủy."
        ],
        signer_authority: "T/L BAN THƯỜNG VỤ",
        signer_title: "CHÁNH VĂN PHÒNG",
        signer_name: "Nguyễn Phước Tâm",
        include_signature: true
    }
};

// Gợi ý nhanh cho Chánh Văn phòng
const QUICK_SUGGESTIONS = {
    leaders: ["TT.ĐU", "BTV", "TT.ĐU, C.Hằng", "A.Trung, A.Hùng", "A.Tùng", "A.Hùng, A.Công", "A.Vũ", "A.Trung, A.Tùng", "A.Vũ, A.Trung(QS)"],
    locations: ["Phòng họp số 1", "Phòng họp số 2", "Phòng họp số 4", "Phòng họp số 5", "Phòng họp BCHQS xã", "Hội trường xã Quế Sơn", "Tại thực tế"],
    staff: ["A.Tâm", "C.Nguyệt", "A.Lực", "A.Lưu, A.Phương", "C.Hằng", "A.Công"],
    commonMeetings: [
        { label: "Họp Thường trực ĐU", time: "08h00", content: "Họp Thường trực Đảng ủy.", leader: "TT.ĐU", location: "Phòng họp số 1" },
        { label: "Họp Ban Thường vụ", time: "14h00", content: "Họp Ban Thường vụ Đảng ủy.", leader: "BTV", location: "Phòng họp số 1", is_red: true },
        { label: "Giao ban UBND", time: "14h00", content: "Họp giao ban lãnh đạo UBND xã.", leader: "A.Trung, A.Tùng", location: "Phòng họp số 1" },
        { label: "Họp HĐND xã", time: "08h00", content: "Họp HĐND xã lần thứ ...", leader: "BTV", location: "Phòng họp số 5" },
        { label: "Sinh hoạt Đảng bộ QS", time: "08h00", content: "Sinh hoạt Đảng bộ Quân sự.", leader: "A.Vũ, A.Trung(QS)", location: "Phòng họp BCHQS xã" },
        { label: "Tiếp công dân / Đi cơ sở", time: "08h00", content: "Đi kiểm tra, khảo sát thực tế tại cơ sở.", leader: "A.Vũ", location: "Tại thực tế" }
    ]
};

// =================== XỬ LÝ NGÀY THÁNG TUẦN ===================

/**
 * Lấy ngày Thứ Hai của tuần ứng với một ngày bất kỳ
 */
function getMonday(d) {
    const date = new Date(d);
    const day = date.getDay();
    const diff = date.getDate() - day + (day === 0 ? -6 : 1); // 0 là Chủ Nhật
    return new Date(date.setDate(diff));
}

/**
 * Định dạng ngày dạng DD/MM/YYYY hoặc DD/MM
 */
function formatDate(d, format = "DD/MM/YYYY") {
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    if (format === "DD/MM") return `${day}/${month}`;
    if (format === "YYYY-MM-DD") return `${year}-${month}-${day}`;
    return `${day}/${month}/${year}`;
}

/**
 * Tính số tuần trong năm (ISO Week)
 */
function getWeekNumber(d) {
    const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

/**
 * Khởi tạo dữ liệu các ngày trong tuần được chọn
 */
function generateWeekDays(monday) {
    const dayNames = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"];
    const days = [];
    for (let i = 0; i < 7; i++) {
        const d = new Date(monday);
        d.setDate(monday.getDate() + i);
        days.push({
            day_name: dayNames[i],
            date_str: formatDate(d, "DD/MM"),
            full_date: formatDate(d, "DD/MM/YYYY"),
            morning: [],
            afternoon: []
        });
    }
    return days;
}

/**
 * Chuyển đổi tuần (0 = Tuần hiện tại, 1 = Tuần kế tiếp/tuần sau...)
 * @param {number} offset - Số tuần lệch so với hiện tại
 * @param {boolean} clearEvents - Có xóa trắng nội dung để soạn mới hay không
 */
function setWeekByOffset(offset, clearEvents = false) {
    appState.currentWeekOffset = offset;
    
    // Ngày hiện tại
    const now = new Date();
    const currentMonday = getMonday(now);
    
    // Tính ngày Thứ Hai của tuần mục tiêu
    const targetMonday = new Date(currentMonday);
    targetMonday.setDate(currentMonday.getDate() + (offset * 7));
    
    // Ngày Chủ Nhật của tuần mục tiêu
    const targetSunday = new Date(targetMonday);
    targetSunday.setDate(targetMonday.getDate() + 6);
    
    // Ngày Thứ Bảy và Chủ Nhật cho ghi chú trực lái xe
    const satDate = new Date(targetMonday);
    satDate.setDate(targetMonday.getDate() + 5);
    
    // Ngày ký thông báo (thường là Chủ nhật trước tuần đó)
    const issueDate = new Date(targetMonday);
    issueDate.setDate(targetMonday.getDate() - 1);
    
    // Số tuần
    const weekNum = getWeekNumber(targetMonday);
    const nextWeekNum = weekNum + 1;
    
    // Cập nhật thông tin tiêu đề và ghi chú
    const fromStr = formatDate(targetMonday);
    const toStr = formatDate(targetSunday);
    const issueDay = String(issueDate.getDate()).padStart(2, '0');
    const issueMonth = String(issueDate.getMonth() + 1).padStart(2, '0');
    appState.data.doc_title = `lịch công tác của Ban Thường vụ Đảng ủy từ ngày ${fromStr} đến ngày ${toStr}`;
    appState.data.doc_date_text = `Quế Sơn, ngày ${issueDay} tháng ${issueMonth} năm ${issueDate.getFullYear()}`;
    appState.data.driver_duty_note = `* Thứ Bảy, Chủ nhật, ngày ${formatDate(satDate, "DD/MM")} – ${formatDate(targetSunday, "DD/MM/YYYY")}. A.Lực trực lái xe Đảng ủy.`;
    appState.data.next_week_title = `* Tuần ${nextWeekNum}:`;
    
    // Cập nhật ngày tháng cho 7 ngày trong bảng
    const dayNames = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"];
    if (clearEvents || !appState.data.days || appState.data.days.length === 0) {
        appState.data.days = generateWeekDays(targetMonday);
        if (clearEvents) {
            appState.data.next_week_items = [];
            appState.data.change_note = "";
        }
    } else {
        // Cập nhật lại nhãn ngày
        for (let i = 0; i < 7; i++) {
            const d = new Date(targetMonday);
            d.setDate(targetMonday.getDate() + i);
            if (!appState.data.days[i]) {
                appState.data.days[i] = {
                    day_name: dayNames[i],
                    date_str: formatDate(d, "DD/MM"),
                    full_date: formatDate(d, "DD/MM/YYYY"),
                    morning: [],
                    afternoon: []
                };
            } else {
                appState.data.days[i].day_name = dayNames[i];
                appState.data.days[i].date_str = formatDate(d, "DD/MM");
                appState.data.days[i].full_date = formatDate(d, "DD/MM/YYYY");
            }
        }
    }
    
    // Cập nhật UI
    updateWeekSelectorUI(targetMonday, targetSunday, weekNum);
    renderEditor();
    renderLivePreview();
}

/**
 * Chuyển chế độ tuần: 'current' (Tuần hiện tại - mặc định) hoặc 'next' (Tuần sau - xóa trắng để lập mới)
 */
function switchWeekMode(mode) {
    const btnCurrent = document.getElementById("btn-week-current");
    const btnNext = document.getElementById("btn-week-next");
    
    if (mode === 'current') {
        setWeekByOffset(0, false);
        if (btnCurrent) {
            btnCurrent.className = "px-3 py-1 text-xs font-semibold rounded-md transition-all whitespace-nowrap bg-brand-600 text-white shadow-sm";
        }
        if (btnNext) {
            btnNext.className = "px-3 py-1 text-xs font-medium rounded-md transition-all whitespace-nowrap text-warm-300 hover:text-white hover:bg-white/10";
        }
    } else if (mode === 'next') {
        setWeekByOffset(1, true);
        if (btnNext) {
            btnNext.className = "px-3 py-1 text-xs font-semibold rounded-md transition-all whitespace-nowrap bg-brand-600 text-white shadow-sm";
        }
        if (btnCurrent) {
            btnCurrent.className = "px-3 py-1 text-xs font-medium rounded-md transition-all whitespace-nowrap text-warm-300 hover:text-white hover:bg-white/10";
        }
        showToast("Đã chuyển sang Tuần sau & xóa trắng nội dung để lập lịch mới!", "info");
    }
}

/**
 * Cập nhật giao diện thanh chọn tuần
 */
function updateWeekSelectorUI(mon, sun, weekNum) {
    const weekLabel = document.getElementById("current-week-label");
    if (weekLabel) {
        weekLabel.textContent = `Tuần ${weekNum} (${formatDate(mon, "DD/MM")} – ${formatDate(sun, "DD/MM/YYYY")})`;
    }
}

// =================== RENDER GIAO DIỆN NHẬP LIỆU ===================

/**
 * Render khu vực biên tập các ngày trong tuần
 */
function renderEditor() {
    const container = document.getElementById("days-editor-container");
    if (!container) return;
    
    container.innerHTML = "";
    
    // Nếu đang ở Tuần hiện tại: Hiển thị thanh trợ giúp điều chỉnh lịch có mốc so sánh
    if (appState.currentWeekOffset === 0) {
        const alertBanner = document.createElement("div");
        alertBanner.className = "bg-amber-50 border border-amber-300 text-amber-950 p-3.5 rounded-xl mb-4 shadow-sm";
        alertBanner.innerHTML = `
            <div class="flex items-start justify-between">
                <div class="flex items-start space-x-2.5">
                    <span class="text-xl leading-none">🔴</span>
                    <div>
                        <div class="font-bold text-xs text-amber-900 flex items-center space-x-1.5">
                            <span>Chế độ: Đang sửa Tuần hiện tại</span>
                            <span class="bg-red-600 text-white text-[10px] px-1.5 py-0.2 rounded font-bold uppercase">Chữ Đỏ = Thay đổi / Bổ sung</span>
                        </div>
                        <div class="text-[11px] text-amber-800 mt-0.5">
                            Quy ước: Các cuộc họp mới thêm hoặc thay đổi giờ/địa điểm/thành phần so với lịch gốc sẽ được <strong>in chữ ĐỎ</strong> để Thường trực & BTV Đảng ủy dễ dàng nhận biết.
                        </div>
                    </div>
                </div>
            </div>
            <div class="mt-2.5 pt-2 border-t border-amber-200 flex flex-wrap items-center justify-between gap-2 text-xs">
                <div class="flex items-center space-x-2">
                    <button type="button" onclick="snapshotBaseSchedule()" 
                            class="px-2.5 py-1 bg-white hover:bg-amber-100 text-amber-900 border border-amber-300 rounded font-semibold transition"
                            title="Lưu lại lịch hiện tại làm mốc ban đầu trước khi sửa">
                        📸 1. Lưu mốc lịch gốc
                    </button>
                    <button type="button" onclick="autoDetectChanges()" 
                            class="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-semibold transition shadow-sm"
                            title="Tự động so sánh với mốc lịch gốc và bật cờ chữ Đỏ cho các mục đã sửa">
                        ⚡ 2. Tự động tô đỏ lịch đã sửa
                    </button>
                </div>
                <label class="inline-flex items-center cursor-pointer select-none text-[11px] text-amber-900 font-medium">
                    <input type="checkbox" ${appState.autoMarkRedOnCurrentWeek ? 'checked' : ''} 
                           onchange="appState.autoMarkRedOnCurrentWeek = this.checked;"
                           class="rounded text-amber-600 mr-1.5">
                    Tự động gắn cờ ĐỎ khi bấm "+ Thêm khung giờ"
                </label>
            </div>
        `;
        container.appendChild(alertBanner);
    }
    
    appState.data.days.forEach((day, dayIndex) => {
        const dayCard = document.createElement("div");
        dayCard.className = "bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mb-4 transition-all";
        
        // Header của ngày
        const totalEvents = (day.morning?.length || 0) + (day.afternoon?.length || 0);
        dayCard.innerHTML = `
            <div class="flex items-center justify-between px-4 py-3 bg-slate-50 border-b border-slate-200 cursor-pointer select-none" onclick="toggleDayCollapse(${dayIndex})">
                <div class="flex items-center space-x-2">
                    <div>
                        <div class="font-bold text-slate-800 text-sm">${day.day_name} <span class="text-blue-600 font-medium">(${day.date_str})</span></div>
                        <div class="text-xs text-slate-500">${totalEvents} công việc đã lên lịch</div>
                    </div>
                </div>
                <div class="flex items-center space-x-2">
                    <button type="button" class="text-xs px-2.5 py-1 bg-white border border-slate-200 rounded-md text-slate-600 hover:bg-slate-100 font-medium" 
                            onclick="event.stopPropagation(); addEvent(${dayIndex}, 'morning')">
                        + Sáng
                    </button>
                    <button type="button" class="text-xs px-2.5 py-1 bg-white border border-slate-200 rounded-md text-slate-600 hover:bg-slate-100 font-medium" 
                            onclick="event.stopPropagation(); addEvent(${dayIndex}, 'afternoon')">
                        + Chiều
                    </button>
                    <svg id="chevron-${dayIndex}" class="w-5 h-5 text-slate-400 transform transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path>
                    </svg>
                </div>
            </div>
            
            <div id="day-content-${dayIndex}" class="p-4 space-y-4">
                <!-- BUỔI SÁNG -->
                <div class="rounded-lg border border-amber-200 bg-amber-50/40 p-3">
                    <div class="flex items-center justify-between mb-2">
                        <div class="flex items-center space-x-2 text-amber-900 font-bold text-xs uppercase tracking-wider">
                            <span>☀️ Buổi Sáng (S)</span>
                            <span class="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 font-normal text-[11px]">${day.morning?.length || 0} mục</span>
                        </div>
                        <button type="button" class="text-xs text-amber-800 hover:text-amber-900 font-semibold flex items-center space-x-1" onclick="addEvent(${dayIndex}, 'morning')">
                            <span>+ Thêm khung giờ</span>
                        </button>
                    </div>
                    <div id="morning-events-${dayIndex}" class="space-y-3">
                        ${renderSessionEvents(day.morning, dayIndex, 'morning')}
                    </div>
                </div>

                <!-- BUỔI CHIỀU -->
                <div class="rounded-lg border border-indigo-200 bg-indigo-50/40 p-3">
                    <div class="flex items-center justify-between mb-2">
                        <div class="flex items-center space-x-2 text-indigo-900 font-bold text-xs uppercase tracking-wider">
                            <span>🌙 Buổi Chiều (C)</span>
                            <span class="px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-800 font-normal text-[11px]">${day.afternoon?.length || 0} mục</span>
                        </div>
                        <button type="button" class="text-xs text-indigo-800 hover:text-indigo-900 font-semibold flex items-center space-x-1" onclick="addEvent(${dayIndex}, 'afternoon')">
                            <span>+ Thêm khung giờ</span>
                        </button>
                    </div>
                    <div id="afternoon-events-${dayIndex}" class="space-y-3">
                        ${renderSessionEvents(day.afternoon, dayIndex, 'afternoon')}
                    </div>
                </div>
            </div>
        `;
        
        container.appendChild(dayCard);
    });
    
    // Đồng bộ các trường thông tin hành chính
    syncMetaInputs();
}

/**
 * Render danh sách các sự kiện trong 1 buổi
 */
function renderSessionEvents(events, dayIndex, session) {
    if (!events || events.length === 0) {
        return `<div class="text-xs text-slate-400 italic text-center py-2 bg-white/60 rounded border border-dashed border-slate-200">
                    Chưa có lịch buổi này. Nhấn "+ Thêm khung giờ" để thêm.
                </div>`;
    }
    
    return events.map((ev, evIndex) => `
        <div class="event-card p-3 rounded-lg border bg-white ${ev.is_red ? 'is-red' : 'border-slate-200'}">
            <div class="grid grid-cols-1 md:grid-cols-12 gap-2 mb-2">
                <!-- TT-BTV ĐU -->
                <div class="md:col-span-3">
                    <label class="block text-[11px] font-semibold text-slate-500 mb-1">Lãnh đạo (TT-BTV ĐU)</label>
                    <input type="text" value="${escapeHtml(ev.leader || '')}" 
                           placeholder="vd: TT.ĐU, BTV..."
                           class="w-full text-xs px-2.5 py-1.5 border rounded border-slate-300 focus:ring-1 focus:ring-blue-500 font-medium ${ev.is_red ? 'text-red-600' : 'text-slate-800'}"
                           oninput="updateEventField(${dayIndex}, '${session}', ${evIndex}, 'leader', this.value)">
                </div>

                <!-- Giờ / Nội dung -->
                <div class="md:col-span-9">
                    <label class="block text-[11px] font-semibold text-slate-500 mb-1">Nội dung công việc / Họp</label>
                    <textarea rows="2" 
                              placeholder="- 08h00: Họp Thường trực Đảng ủy..."
                              class="w-full text-xs px-2.5 py-1.5 border rounded border-slate-300 focus:ring-1 focus:ring-blue-500 ${ev.is_red ? 'text-red-600 font-semibold' : 'text-slate-800'}"
                              oninput="updateEventField(${dayIndex}, '${session}', ${evIndex}, 'content', this.value)">${escapeHtml(ev.content || '')}</textarea>
                </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-12 gap-2 mb-2">
                <!-- Địa điểm -->
                <div class="md:col-span-6">
                    <label class="block text-[11px] font-semibold text-slate-500 mb-1">Địa điểm</label>
                    <input type="text" value="${escapeHtml(ev.location || '')}" 
                           placeholder="vd: Phòng họp số 1..."
                           class="w-full text-xs px-2.5 py-1.5 border rounded border-slate-300 focus:ring-1 focus:ring-blue-500 ${ev.is_red ? 'text-red-600' : 'text-slate-800'}"
                           oninput="updateEventField(${dayIndex}, '${session}', ${evIndex}, 'location', this.value)">
                </div>

                <!-- VP + LX -->
                <div class="md:col-span-6">
                    <label class="block text-[11px] font-semibold text-slate-500 mb-1">VP + LX (Cán bộ đi cùng)</label>
                    <input type="text" value="${escapeHtml(ev.staff || '')}" 
                           placeholder="vd: A.Tâm, C.Nguyệt..."
                           class="w-full text-xs px-2.5 py-1.5 border rounded border-slate-300 focus:ring-1 focus:ring-blue-500 ${ev.is_red ? 'text-red-600' : 'text-slate-800'}"
                           oninput="updateEventField(${dayIndex}, '${session}', ${evIndex}, 'staff', this.value)">
                </div>
            </div>

            <!-- Toolbar phụ: Đánh dấu đỏ, Gợi ý nhanh, Xóa, Nhân bản -->
            <div class="flex flex-wrap items-center justify-between pt-2 border-t border-slate-100 gap-2">
                <div class="flex items-center space-x-2">
                    <label class="inline-flex items-center cursor-pointer select-none">
                        <input type="checkbox" ${ev.is_red ? 'checked' : ''} 
                               class="rounded border-slate-300 text-red-600 focus:ring-red-500"
                               onchange="updateEventField(${dayIndex}, '${session}', ${evIndex}, 'is_red', this.checked)">
                        <span class="ml-1.5 text-xs font-semibold ${ev.is_red ? 'text-red-600' : 'text-slate-600'} flex items-center space-x-1">
                            <span>🔴 Lịch thay đổi / bổ sung so với lịch gốc (In chữ ĐỎ)</span>
                            ${ev.is_red ? '<span class="ml-1 text-[10px] bg-red-100 text-red-800 px-1 py-0.2 rounded font-bold uppercase">ĐÃ ĐỔI</span>' : ''}
                        </span>
                    </label>
                </div>

                <div class="flex items-center space-x-1">
                    <button type="button" class="text-xs text-slate-500 hover:text-blue-600 px-2 py-1 rounded hover:bg-slate-100" 
                            title="Nhân bản sự kiện này"
                            onclick="duplicateEvent(${dayIndex}, '${session}', ${evIndex})">
                        📋 Nhân bản
                    </button>
                    <button type="button" class="text-xs text-red-500 hover:text-red-700 px-2 py-1 rounded hover:bg-red-50" 
                            title="Xóa sự kiện"
                            onclick="deleteEvent(${dayIndex}, '${session}', ${evIndex})">
                        🗑️ Xóa
                    </button>
                </div>
            </div>
        </div>
    `).join("");
}

/**
 * Đồng bộ các trường thông tin chung (Cơ quan, ngày tháng, ghi chú, người ký...)
 */
function syncMetaInputs() {
    const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val || "";
    };
    
    setVal("input-agency-upper", appState.data.agency_upper);
    setVal("input-agency-name", appState.data.agency_name);
    setVal("input-doc-date", appState.data.doc_date_text);
    setVal("input-doc-title", appState.data.doc_title);
    setVal("input-driver-note", appState.data.driver_duty_note);
    setVal("input-invitation-note", appState.data.invitation_note);
    setVal("input-change-note", appState.data.change_note);
    setVal("input-next-week-title", appState.data.next_week_title);
    setVal("input-signer-authority", appState.data.signer_authority);
    setVal("input-signer-title", appState.data.signer_title);
    setVal("input-signer-name", appState.data.signer_name);
    
    // Nơi nhận
    const recText = (appState.data.recipients || []).join("\n");
    setVal("input-recipients", recText);
    
    // Dự kiến tuần sau
    const nwText = (appState.data.next_week_items || []).join("\n");
    setVal("input-next-week-items", nwText);
    
    // Checkbox chữ ký
    const chkSig = document.getElementById("check-include-signature");
    if (chkSig) chkSig.checked = appState.data.include_signature !== false;
}

// =================== CÁC THAO TÁC SỰ KIỆN ===================

function toggleDayCollapse(dayIndex) {
    const el = document.getElementById(`day-content-${dayIndex}`);
    const chevron = document.getElementById(`chevron-${dayIndex}`);
    if (el) {
        el.classList.toggle("hidden");
        if (chevron) chevron.classList.toggle("rotate-180");
    }
}

/**
 * Lưu mốc lịch gốc ban đầu (Snapshot) để so sánh thay đổi
 */
function snapshotBaseSchedule() {
    appState.baseSnapshot = JSON.parse(JSON.stringify(appState.data.days));
    showToast("Đã lưu mốc Lịch gốc ban đầu! Giờ bạn có thể sửa và hệ thống sẽ tự động nhận diện lịch thay đổi.", "success");
    renderEditor();
}

/**
 * Tự động so sánh với mốc lịch gốc để tô đỏ tất cả nội dung đã thay đổi / bổ sung
 */
function autoDetectChanges() {
    if (!appState.baseSnapshot) {
        showToast("Chưa có mốc lịch gốc. Hãy nhấn 'Lưu mốc lịch gốc' trước khi chỉnh sửa.", "info");
        return;
    }
    
    let changeCount = 0;
    appState.data.days.forEach((day, dIdx) => {
        const baseDay = appState.baseSnapshot[dIdx];
        if (!baseDay) return;
        
        ['morning', 'afternoon'].forEach(session => {
            const currentEvents = day[session] || [];
            const baseEvents = baseDay[session] || [];
            
            currentEvents.forEach((ev, evIdx) => {
                const baseEv = baseEvents[evIdx];
                if (!baseEv) {
                    // Sự kiện mới thêm vào -> Lịch bổ sung
                    ev.is_red = true;
                    changeCount++;
                } else {
                    // So sánh nội dung, giờ, người, địa điểm
                    const isDifferent = (
                        (ev.leader || '').trim() !== (baseEv.leader || '').trim() ||
                        (ev.content || '').trim() !== (baseEv.content || '').trim() ||
                        (ev.location || '').trim() !== (baseEv.location || '').trim() ||
                        (ev.staff || '').trim() !== (baseEv.staff || '').trim()
                    );
                    if (isDifferent) {
                        ev.is_red = true;
                        changeCount++;
                    }
                }
            });
        });
    });
    
    renderEditor();
    renderLivePreview();
    showToast(`Đã tự động nhận diện và tô đỏ ${changeCount} mục lịch có thay đổi/bổ sung!`, "success");
}

function addEvent(dayIndex, session) {
    if (!appState.data.days[dayIndex][session]) {
        appState.data.days[dayIndex][session] = [];
    }
    const defaultTime = session === "morning" ? "08h00" : "14h00";
    // Nếu đang ở Tuần hiện tại thì mặc định đánh dấu là Lịch thay đổi/bổ sung (chữ Đỏ)
    const shouldMarkRed = (appState.currentWeekOffset === 0 && appState.autoMarkRedOnCurrentWeek);
    
    appState.data.days[dayIndex][session].push({
        leader: "",
        content: `- ${defaultTime}: `,
        location: "Phòng họp số 1",
        staff: "",
        is_red: shouldMarkRed
    });
    renderEditor();
    renderLivePreview();
}

function deleteEvent(dayIndex, session, evIndex) {
    if (confirm("Bạn có chắc muốn xóa lịch này không?")) {
        appState.data.days[dayIndex][session].splice(evIndex, 1);
        renderEditor();
        renderLivePreview();
    }
}

function duplicateEvent(dayIndex, session, evIndex) {
    const original = appState.data.days[dayIndex][session][evIndex];
    const clone = JSON.parse(JSON.stringify(original));
    appState.data.days[dayIndex][session].splice(evIndex + 1, 0, clone);
    renderEditor();
    renderLivePreview();
}

function updateEventField(dayIndex, session, evIndex, field, value) {
    const ev = appState.data.days[dayIndex][session][evIndex];
    if (ev) {
        ev[field] = value;
        // Cập nhật ngay Live Preview
        renderLivePreview();
    }
}

function onMetaFieldChange(field, value) {
    appState.data[field] = value;
    renderLivePreview();
}

function onRecipientsChange(value) {
    appState.data.recipients = value.split("\n").filter(s => s.trim().length > 0);
    renderLivePreview();
}

function onNextWeekItemsChange(value) {
    appState.data.next_week_items = value.split("\n").filter(s => s.trim().length > 0);
    renderLivePreview();
}

// =================== LIVE PREVIEW RENDERER ===================

/**
 * Render bản xem trước trực quan A4 chuẩn 1:1
 */
function renderLivePreview() {
    const previewContainer = document.getElementById("a4-preview-content");
    if (!previewContainer) return;
    
    const d = appState.data;
    
    // Xây dựng các hàng của bảng
    let tableRowsHtml = "";
    
    (d.days || []).forEach((day) => {
        const morningEvents = day.morning || [];
        const afternoonEvents = day.afternoon || [];
        const isWeekend = /thứ bảy|thứ 7|chủ nhật/i.test(day.day_name || "");
        
        const hasContent = morningEvents.some(e => (e.content || '').trim() || (e.leader || '').trim()) ||
                           afternoonEvents.some(e => (e.content || '').trim() || (e.leader || '').trim());
                           
        // Bỏ qua Thứ 7 và Chủ nhật nếu không có nội dung công việc
        if (isWeekend && !hasContent && morningEvents.length === 0 && afternoonEvents.length === 0) {
            return;
        }

        const morningCount = morningEvents.length;
        const afternoonCount = afternoonEvents.length;
        const totalRows = (morningCount + afternoonCount) || 1;
        
        let dayRendered = false;
        let morningRendered = false;
        let afternoonRendered = false;
        
        // Buổi sáng
        if (morningCount > 0) {
            morningEvents.forEach((ev, evIndex) => {
                const isLastInSession = (evIndex === morningCount - 1);
                const borderClasses = [];
                if (!isLastInSession) borderClasses.push("cell-dotted-bottom");
                if (evIndex > 0) borderClasses.push("cell-dotted-top");
                const borderClassStr = borderClasses.join(" ");
                const redClass = ev.is_red ? "text-red-important font-semibold" : "";
                tableRowsHtml += `<tr>`;
                
                // Cột Ngày (rowspan cho cả ngày)
                if (!dayRendered) {
                    tableRowsHtml += `
                        <td rowspan="${totalRows}" style="width: 12%; text-align: center; font-weight: bold; vertical-align: middle;">
                            ${escapeHtml(day.day_name)}<br>${escapeHtml(day.date_str)}
                        </td>`;
                    dayRendered = true;
                }
                
                // Cột Buổi Sáng (rowspan cho buổi sáng)
                if (!morningRendered) {
                    tableRowsHtml += `
                        <td rowspan="${morningCount}" style="width: 5%; text-align: center; font-weight: bold; vertical-align: middle;">
                            S
                        </td>`;
                    morningRendered = true;
                }
                
                // Các cột còn lại
                tableRowsHtml += `
                    <td class="${redClass} ${borderClassStr}" style="width: 16%; text-align: center; vertical-align: middle; white-space: pre-line;">${escapeHtml(ev.leader || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 42%; text-align: left; vertical-align: middle; line-height: 1.25;">${escapeHtml(ev.content || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 14%; text-align: center; vertical-align: middle; line-height: 1.25;">${escapeHtml(ev.location || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 11%; text-align: center; vertical-align: middle; font-size: 11pt; line-height: 1.25; white-space: pre-line;">${escapeHtml(ev.staff || '')}</td>
                </tr>`;
            });
        }
        
        // Buổi chiều
        if (afternoonCount > 0) {
            afternoonEvents.forEach((ev, evIndex) => {
                const isLastInSession = (evIndex === afternoonCount - 1);
                const borderClasses = [];
                if (!isLastInSession) borderClasses.push("cell-dotted-bottom");
                if (evIndex > 0) borderClasses.push("cell-dotted-top");
                const borderClassStr = borderClasses.join(" ");
                const redClass = ev.is_red ? "text-red-important font-semibold" : "";
                tableRowsHtml += `<tr>`;
                
                // Nếu chưa render cột Ngày (trong trường hợp buổi sáng không có gì)
                if (!dayRendered) {
                    tableRowsHtml += `
                        <td rowspan="${totalRows}" style="width: 12%; text-align: center; font-weight: bold; vertical-align: middle;">
                            ${escapeHtml(day.day_name)}<br>${escapeHtml(day.date_str)}
                        </td>`;
                    dayRendered = true;
                }
                
                // Cột Buổi Chiều (rowspan cho buổi chiều)
                if (!afternoonRendered) {
                    tableRowsHtml += `
                        <td rowspan="${afternoonCount}" style="width: 5%; text-align: center; font-weight: bold; vertical-align: middle;">
                            C
                        </td>`;
                    afternoonRendered = true;
                }
                
                tableRowsHtml += `
                    <td class="${redClass} ${borderClassStr}" style="width: 16%; text-align: center; vertical-align: middle; white-space: pre-line;">${escapeHtml(ev.leader || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 42%; text-align: left; vertical-align: middle; line-height: 1.25;">${escapeHtml(ev.content || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 14%; text-align: center; vertical-align: middle; line-height: 1.25;">${escapeHtml(ev.location || '')}</td>
                    <td class="${redClass} ${borderClassStr}" style="width: 11%; text-align: center; vertical-align: middle; font-size: 11pt; line-height: 1.25; white-space: pre-line;">${escapeHtml(ev.staff || '')}</td>
                </tr>`;
            });
        }
        
        // Trường hợp cả sáng và chiều đều trống (cho các ngày Thứ Hai - Thứ Sáu)
        if (morningCount === 0 && afternoonCount === 0) {
            tableRowsHtml += `<tr>
                <td style="width: 12%; text-align: center; font-weight: bold; vertical-align: middle;">
                    ${escapeHtml(day.day_name)}<br>${escapeHtml(day.date_str)}
                </td>
                <td style="width: 5%; text-align: center; font-weight: bold; vertical-align: middle;"></td>
                <td style="width: 16%;"></td>
                <td style="width: 42%; text-align: center; color: #94a3b8; font-style: italic;">(Không có lịch)</td>
                <td style="width: 14%;"></td>
                <td style="width: 11%;"></td>
            </tr>`;
        }
    });

    // Phần Nơi nhận (Không in đậm, ngang hàng lề trái)
    const recipientsHtml = (d.recipients || []).map(r => `<div style="text-align: left; margin: 0; padding: 0; line-height: 1.25;">- ${escapeHtml(r.replace(/^-\s*/, ''))}</div>`).join("");

    // Phần Dự kiến tuần sau
    const nextWeekHtml = (d.next_week_items || []).map(it => `<div style="padding-left: 12px; margin-bottom: 3px; font-style: italic;">${escapeHtml(it)}</div>`).join("");

    // Phần Chữ ký
    const sigImgHtml = d.include_signature ? `
        <div style="display: flex; justify-content: center; margin: 4px 0;">
            <img src="/static/signature.png" alt="Chữ ký" style="height: 55px; object-fit: contain;">
        </div>` : `<div style="height: 60px;"></div>`;

    previewContainer.innerHTML = `
        <!-- TIÊU NGỮ 2 CỘT -->
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
            <tr>
                <td style="width: 48%; text-align: center; vertical-align: top; font-size: 12.5pt; line-height: 1.2;">
                    <div>${escapeHtml(d.agency_upper || '')}</div>
                    <div style="font-weight: bold; font-size: 13pt;">${escapeHtml(d.agency_name || '')}</div>
                    <div style="font-weight: bold; margin-top: 2px;">*</div>
                </td>
                <td style="width: 52%; text-align: center; vertical-align: top; line-height: 1.25;">
                    <div style="font-weight: bold; text-decoration: underline; font-size: 13.5pt;">ĐẢNG CỘNG SẢN VIỆT NAM</div>
                    <div style="font-style: italic; font-size: 13pt; text-align: center; margin-top: 5px;">
                        ${escapeHtml(d.doc_date_text || '')}
                    </div>
                </td>
            </tr>
        </table>

        <!-- TIÊU ĐỀ THÔNG BÁO -->
        <div style="text-align: center; margin: 10px 0 16px 0;">
            <div style="font-size: 15pt; font-weight: bold; letter-spacing: 0.5px;">THÔNG BÁO</div>
            <div style="font-size: 13pt; font-weight: bold; margin-top: 2px;">${escapeHtml(d.doc_title || '')}</div>
            <div style="font-weight: bold; font-size: 12pt; margin-top: 2px;">-----</div>
        </div>

        <!-- BẢNG LỊCH CÔNG TÁC -->
        <table class="preview-table">
            <thead>
                <tr>
                    <th colspan="2" style="width: 17%; font-size: 13pt; padding: 5px;">Ngày</th>
                    <th style="width: 16%; font-size: 12pt; padding: 5px;">TT-BTV ĐU</th>
                    <th style="width: 42%; font-size: 12pt; padding: 5px;">Nội dung</th>
                    <th style="width: 14%; font-size: 12pt; padding: 5px;">Địa điểm</th>
                    <th style="width: 11%; font-size: 11pt; padding: 5px;">VP + LX</th>
                </tr>
            </thead>
            <tbody>
                ${tableRowsHtml}
            </tbody>
        </table>

        <!-- GHI CHÚ CUỐI TUẦN & DỰ KIẾN TUẦN SAU -->
        <div style="margin-top: 10px; font-size: 12pt; line-height: 1.3;">
            <div style="font-style: italic; margin-bottom: 3px;">${escapeHtml(d.driver_duty_note || '')}</div>
            <div style="font-style: italic; margin-bottom: 4px;">${escapeHtml(d.invitation_note || '')}</div>
            ${d.change_note ? `<div style="font-style: italic; color: #dc2626; margin-bottom: 4px;">${escapeHtml(d.change_note)}</div>` : ''}
            
            ${d.next_week_title ? `<div style="font-weight: bold; font-style: italic; margin-top: 6px; margin-bottom: 4px;">${escapeHtml(d.next_week_title)}</div>` : ''}
            ${nextWeekHtml}
        </div>

        <!-- NƠI NHẬN & CHỮ KÝ -->
        <table style="width: 100%; border-collapse: collapse; margin-top: 18px;">
            <tr>
                <td style="width: 45%; vertical-align: top; font-size: 11pt; line-height: 1.25; padding: 0; text-align: left;">
                    <div style="font-weight: normal; text-decoration: underline; font-size: 12pt; margin-bottom: 4px; text-align: left;">Nơi nhận:</div>
                    ${recipientsHtml}
                </td>
                <td style="width: 55%; text-align: center; vertical-align: top; font-size: 12.5pt; line-height: 1.25;">
                    <div style="font-weight: bold;">${escapeHtml(d.signer_authority || '')}</div>
                    <div style="font-weight: normal;">${escapeHtml(d.signer_title || '')}</div>
                    ${sigImgHtml}
                    <div style="font-weight: bold; font-size: 13pt;">${escapeHtml(d.signer_name || '')}</div>
                </td>
            </tr>
        </table>
    `;
}

// =================== XUẤT FILE WORD (.DOCX) ===================

function getStandardExportFilename(d) {
    const docTitle = d.doc_title || "";
    let weekNum = null;
    const nextWeekTitle = d.next_week_title || "";
    const mNw = nextWeekTitle.match(/Tuần\s*(\d+)/i);
    if (mNw) {
        weekNum = parseInt(mNw[1]) - 1;
    }
    
    const dateMatches = [...docTitle.matchAll(/(\d{1,2})[/-](\d{1,2})(?:[/-](\d{4}))?/g)];
    let startStr = "";
    let endStr = "";
    const currentYear = new Date().getFullYear();
    
    if (dateMatches.length >= 2) {
        const d1 = dateMatches[0][1].padStart(2, '0');
        const m1 = dateMatches[0][2].padStart(2, '0');
        const d2 = dateMatches[1][1].padStart(2, '0');
        const m2 = dateMatches[1][2].padStart(2, '0');
        const y2 = dateMatches[1][3] || dateMatches[0][3] || currentYear;
        startStr = `${d1}.${m1}`;
        endStr = `${d2}.${m2}.${y2}`;
    } else if (d.days && d.days.length > 0) {
        const fd = d.days[0].date_str.replace('/', '.');
        const ld = d.days[d.days.length - 1].date_str.replace('/', '.');
        startStr = fd;
        endStr = `${ld}.${currentYear}`;
    }
    
    if (!weekNum) {
        const weekLabel = document.getElementById("current-week-label");
        const mLabel = weekLabel ? weekLabel.textContent.match(/Tuần\s*(\d+)/i) : null;
        weekNum = mLabel ? parseInt(mLabel[1]) : 45;
    }
    
    if (!startStr) startStr = "02.11";
    if (!endStr) endStr = `08.11.${currentYear}`;
    
    return `Tuần ${weekNum} lịch công tác BTV ${startStr} - ${endStr}.docx`;
}

async function exportDocx() {
    const btn = document.getElementById("btn-export-docx");
    const originalText = btn.innerHTML;
    
    try {
        btn.disabled = true;
        btn.innerHTML = `
            <svg class="animate-spin -ml-1 mr-2 h-4 w-4 text-white inline-block" fill="none" viewBox="0 0 24 24">
                <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
            </svg>
            Đang tạo file Word...
        `;
        
        const response = await fetch("/api/export-docx", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(appState.data)
        });
        
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || "Lỗi khi xuất file");
        }
        
        const blob = await response.blob();
        
        // Lấy tên file xuất chuẩn xác theo cấu trúc "Tuần 45 lịch công tác BTV 02.11 - 08.11.2026.docx"
        let filename = "";
        const xFilename = response.headers.get("X-Download-Filename");
        if (xFilename) {
            try {
                filename = decodeURIComponent(xFilename);
            } catch (e) {
                filename = xFilename;
            }
        }
        
        if (!filename) {
            filename = getStandardExportFilename(appState.data);
        }
        
        // Tải xuống
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
        
        showToast(`Tải file "${filename}" thành công! Đúng 100% mẫu công văn Đảng.`, "success");
    } catch (e) {
        console.error(e);
        showToast("Lỗi khi xuất file Word: " + e.message, "error");
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
}

// =================== QUẢN LÝ KHO LỊCH CŨ (HISTORY) ===================

let cachedHistoryList = [];

/**
 * Mở hộp thoại Kho Lịch Cũ
 */
async function openHistoryModal() {
    const modal = document.getElementById("history-modal");
    if (!modal) return;
    modal.classList.remove("hidden");
    await fetchAndRenderHistory();
}

/**
 * Đóng hộp thoại
 */
function closeHistoryModal() {
    const modal = document.getElementById("history-modal");
    if (modal) modal.classList.add("hidden");
}

/**
 * Lấy danh sách lịch cũ từ Server và hiển thị
 */
async function fetchAndRenderHistory() {
    const container = document.getElementById("history-list-container");
    if (!container) return;
    
    container.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs">Đang tải danh sách lịch cũ...</div>`;
    
    try {
        const res = await fetch("/api/history");
        const list = await res.json();
        cachedHistoryList = list;
        
        // Cập nhật số lượng trên badge
        const badge = document.getElementById("history-badge-count");
        if (badge) badge.textContent = list.length;
        
        renderHistoryList(list);
    } catch (e) {
        container.innerHTML = `<div class="text-center py-8 text-red-500 text-xs">Lỗi tải danh sách: ${e.message}</div>`;
    }
}

/**
 * Render danh sách lịch cũ ra giao diện
 */
function renderHistoryList(list) {
    const container = document.getElementById("history-list-container");
    if (!container) return;
    
    if (!list || list.length === 0) {
        container.innerHTML = `
            <div class="text-center py-10 bg-slate-50 rounded-xl border border-dashed border-slate-200">
                <div class="text-3xl mb-2">📭</div>
                <div class="font-bold text-slate-700 text-xs">Chưa có bản lịch nào trong kho</div>
                <div class="text-slate-400 text-[11px] mt-1">Hãy bấm "Lưu lịch" ở thanh công cụ hoặc nhập file .json để lưu trữ.</div>
            </div>`;
        return;
    }
    
    container.innerHTML = list.map(item => `
        <div class="bg-white border border-slate-200 hover:border-blue-400 rounded-xl p-4 shadow-sm hover:shadow transition-all">
            <div class="flex items-start justify-between gap-3">
                <div class="flex-1">
                    <div class="flex items-center space-x-2">
                        <span class="font-bold text-slate-900 text-sm hover:text-blue-600 transition cursor-pointer" 
                              onclick="loadHistoryItem('${item.id}', 'current')">
                            ${escapeHtml(item.name || item.doc_title)}
                        </span>
                        ${item.red_events > 0 ? `<span class="bg-red-100 text-red-700 font-bold text-[10px] px-1.5 py-0.5 rounded uppercase">Có ${item.red_events} lịch đổi</span>` : ''}
                    </div>
                    
                    <div class="text-[11px] text-slate-500 mt-1 flex flex-wrap items-center gap-x-3 gap-y-1">
                        <span>📅 ${escapeHtml(item.doc_date_text || item.doc_title)}</span>
                        <span>•</span>
                        <span>📝 ${item.total_events} cuộc họp</span>
                        <span>•</span>
                        <span>⏰ Đã lưu: ${escapeHtml(item.saved_at)}</span>
                    </div>
                </div>

                <div class="flex items-center space-x-1 flex-shrink-0">
                    <button type="button" onclick="loadHistoryItem('${item.id}', 'current')" 
                            class="text-xs px-2.5 py-1.5 bg-brand-50 hover:bg-brand-600 hover:text-white text-brand-700 font-bold rounded-xl border border-brand-200 transition flex items-center space-x-1"
                            title="Lấy nội dung lịch cũ này nạp vào tuần bạn đang chọn trên màn hình">
                        <span>⬇ Nạp sang tuần đang chọn</span>
                    </button>

                    <button type="button" onclick="loadHistoryItem('${item.id}', 'exact')" 
                            class="text-xs px-2 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-lg transition"
                            title="Mở nguyên bản gốc (giữ nguyên ngày tháng cũ)">
                        <span>📖 Mở gốc</span>
                    </button>

                    <button type="button" onclick="deleteHistoryItem('${item.id}')" 
                            class="text-xs p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition"
                            title="Xóa lịch này khỏi kho">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                        </svg>
                    </button>
                </div>
            </div>
        </div>
    `).join("");
}

/**
 * Tìm kiếm trong danh sách lịch cũ
 */
function filterHistoryList(query) {
    const q = (query || "").toLowerCase().trim();
    if (!q) {
        renderHistoryList(cachedHistoryList);
        return;
    }
    const filtered = cachedHistoryList.filter(it => 
        (it.name || '').toLowerCase().includes(q) ||
        (it.doc_title || '').toLowerCase().includes(q) ||
        (it.doc_date_text || '').toLowerCase().includes(q) ||
        (it.saved_at || '').toLowerCase().includes(q)
    );
    renderHistoryList(filtered);
}

/**
 * Nạp một bản lịch cũ
 * @param {string} itemId
 * @param {'current' | 'exact'} mode 
 * - 'current': Nạp nội dung các sự kiện vào tuần đang chọn (giữ nguyên ngày tháng của tuần mới)
 * - 'exact': Nạp nguyên trạng cả ngày tháng ban đầu của lịch cũ
 */
async function loadHistoryItem(itemId, mode = 'current') {
    try {
        const res = await fetch(`/api/history/${itemId}`);
        if (!res.ok) throw new Error("Không thể tải lịch này");
        const loadedData = await res.json();
        
        if (mode === 'current') {
            // Lưu lại thông tin ngày tháng của tuần đang chọn
            const currentDays = appState.data.days || [];
            
            // Sao chép cơ quan, ghi chú, người ký
            appState.data.agency_upper = loadedData.agency_upper || appState.data.agency_upper;
            appState.data.agency_name = loadedData.agency_name || appState.data.agency_name;
            appState.data.signer_authority = loadedData.signer_authority || appState.data.signer_authority;
            appState.data.signer_title = loadedData.signer_title || appState.data.signer_title;
            appState.data.signer_name = loadedData.signer_name || appState.data.signer_name;
            appState.data.recipients = loadedData.recipients || appState.data.recipients;
            
            // Ghép nội dung các sự kiện vào các ngày của tuần đang chọn
            for (let i = 0; i < 7; i++) {
                if (currentDays[i] && loadedData.days && loadedData.days[i]) {
                    currentDays[i].morning = JSON.parse(JSON.stringify(loadedData.days[i].morning || []));
                    currentDays[i].afternoon = JSON.parse(JSON.stringify(loadedData.days[i].afternoon || []));
                }
            }
            appState.data.days = currentDays;
            showToast(`Đã nạp lịch "${loadedData.name || loadedData.doc_title}" sang tuần đang chọn!`, "success");
        } else {
            // Nạp nguyên bản gốc
            appState.data = loadedData;
            showToast(`Đã mở nguyên bản gốc của "${loadedData.name || loadedData.doc_title}"!`, "success");
        }
        
        closeHistoryModal();
        renderEditor();
        renderLivePreview();
    } catch (e) {
        showToast("Lỗi nạp lịch: " + e.message, "error");
    }
}

/**
 * Lưu bản lịch hiện tại vào kho lịch cũ
 */
async function saveCurrentToHistory() {
    const defaultName = appState.data.doc_title || `Lịch công tác ${formatDate(new Date())}`;
    const name = prompt("Nhập tên để lưu lịch này vào kho lưu trữ:", defaultName);
    if (!name) return;
    
    try {
        const payload = Object.assign({}, appState.data, { save_name: name });
        const res = await fetch("/api/history/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });
        
        if (!res.ok) throw new Error("Không thể lưu lịch");
        const resJson = await res.json();
        
        showToast(`Đã lưu thành công lịch "${resJson.name}" vào kho!`, "success");
        
        // Cập nhật lại số lượng badge
        const histRes = await fetch("/api/history");
        const list = await histRes.json();
        const badge = document.getElementById("history-badge-count");
        if (badge) badge.textContent = list.length;
    } catch (e) {
        showToast("Lỗi khi lưu lịch: " + e.message, "error");
    }
}

/**
 * Xóa một lịch trong kho
 */
async function deleteHistoryItem(itemId) {
    if (!confirm("Bạn có chắc chắn muốn xóa bản lịch này khỏi kho không?")) return;
    
    try {
        const res = await fetch(`/api/history/${itemId}`, { method: "DELETE" });
        if (!res.ok) throw new Error("Không thể xóa lịch này");
        showToast("Đã xóa bản lịch thành công!", "info");
        await fetchAndRenderHistory();
    } catch (e) {
        showToast("Lỗi xóa lịch: " + e.message, "error");
    }
}

/**
 * Xử lý nạp lịch trực tiếp từ file Word (.docx hoặc .doc) đã xuất tuần trước
 */
async function handleWordFileSelected(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.docx' && ext !== '.doc') {
        showToast("Vui lòng chọn file Word có đuôi .docx hoặc .doc", "error");
        event.target.value = "";
        return;
    }
    
    const formData = new FormData();
    formData.append("file", file);
    
    showToast(`Đang đọc và bóc tách dữ liệu từ file "${file.name}"...`, "info");
    
    try {
        const response = await fetch("/api/parse-word", {
            method: "POST",
            body: formData
        });
        
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || "Không thể phân tích file Word này");
        }
        
        const resJson = await response.json();
        const loadedData = resJson.data;
        
        // Đếm số lượng cuộc họp bóc tách được
        let totalMeetings = 0;
        let redMeetings = 0;
        (loadedData.days || []).forEach(d => {
            (d.morning || []).forEach(ev => { totalMeetings++; if (ev.is_red) redMeetings++; });
            (d.afternoon || []).forEach(ev => { totalMeetings++; if (ev.is_red) redMeetings++; });
        });
        
        // Hỏi người dùng phương thức nạp
        const applyToCurrent = confirm(
            `Đã bóc tách thành công ${totalMeetings} cuộc họp (${redMeetings} lịch đổi) từ file "${file.name}"!\n\n` +
            `• Bấm [OK]: Nạp nội dung lịch này vào TUẦN ĐANG CHỌN (giữ nguyên ngày tháng của tuần mới).\n` +
            `• Bấm [Cancel]: Mở NGUYÊN BẢN GỐC (giữ nguyên ngày tháng cũ của file Word).`
        );
        
        if (applyToCurrent) {
            // Nạp nội dung vào tuần đang chọn
            const currentDays = appState.data.days || [];
            
            appState.data.agency_upper = loadedData.agency_upper || appState.data.agency_upper;
            appState.data.agency_name = loadedData.agency_name || appState.data.agency_name;
            appState.data.signer_authority = loadedData.signer_authority || appState.data.signer_authority;
            appState.data.signer_title = loadedData.signer_title || appState.data.signer_title;
            appState.data.signer_name = loadedData.signer_name || appState.data.signer_name;
            appState.data.recipients = loadedData.recipients || appState.data.recipients;
            
            for (let i = 0; i < 7; i++) {
                if (currentDays[i] && loadedData.days && loadedData.days[i]) {
                    currentDays[i].morning = JSON.parse(JSON.stringify(loadedData.days[i].morning || []));
                    currentDays[i].afternoon = JSON.parse(JSON.stringify(loadedData.days[i].afternoon || []));
                }
            }
            appState.data.days = currentDays;
            showToast(`Đã áp dụng ${totalMeetings} cuộc họp từ file Word sang tuần đang chọn!`, "success");
        } else {
            // Mở nguyên bản gốc
            appState.data = loadedData;
            showToast(`Đã mở nguyên bản gốc từ file Word "${file.name}"!`, "success");
        }
        
        closeHistoryModal();
        renderEditor();
        renderLivePreview();
        
        // Cập nhật lại số lượng badge
        fetch("/api/history")
            .then(r => r.json())
            .then(list => {
                const badge = document.getElementById("history-badge-count");
                if (badge && Array.isArray(list)) badge.textContent = list.length;
            });
            
    } catch (e) {
        console.error(e);
        showToast("Lỗi khi đọc file Word: " + e.message, "error");
    } finally {
        event.target.value = "";
    }
}

/**
 * Xử lý nhập file .json từ máy tính
 */
function handleFileImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const json = JSON.parse(e.target.result);
            if (!json.days || !Array.isArray(json.days)) {
                throw new Error("File không đúng cấu trúc lịch tuần");
            }
            
            // Tự động lưu vào kho
            const saveName = json.name || json.doc_title || file.name.replace(".json", "");
            json.save_name = saveName;
            
            await fetch("/api/history/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(json)
            });
            
            showToast(`Đã nhập file "${file.name}" vào kho thành công!`, "success");
            await fetchAndRenderHistory();
        } catch (err) {
            showToast("Lỗi đọc file JSON: " + err.message, "error");
        }
    };
    reader.readAsText(file);
    // Reset file input
    event.target.value = "";
}

function clearAllEvents() {
    if (confirm("Bạn có chắc chắn muốn xóa toàn bộ lịch các ngày để nhập mới không?")) {
        (appState.data.days || []).forEach(d => {
            d.morning = [];
            d.afternoon = [];
        });
        renderEditor();
        renderLivePreview();
        showToast("Đã xóa trắng lịch tuần để sẵn sàng nhập mới.", "info");
    }
}

function printDocument() {
    window.print();
}

function setZoom(level) {
    appState.zoomLevel = level;
    const page = document.querySelector(".a4-page");
    if (page) {
        page.style.transform = `scale(${level})`;
    }
    const zoomText = document.getElementById("zoom-level-text");
    if (zoomText) zoomText.textContent = `${Math.round(level * 100)}%`;
}

function showToast(message, type = "info") {
    const toast = document.createElement("div");
    const colors = {
        success: "bg-emerald-600 text-white",
        error: "bg-red-600 text-white",
        info: "bg-brand-600 text-white"
    };
    toast.className = `fixed bottom-6 right-6 px-4 py-3 rounded-lg shadow-xl ${colors[type] || colors.info} z-50 flex items-center space-x-2 animate-fade-in font-medium text-sm`;
    toast.innerHTML = `<span>${message}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = "0";
        toast.style.transition = "opacity 0.3s ease";
        setTimeout(() => toast.remove(), 300);
    }, 3500);
}

function escapeHtml(str) {
    if (!str) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function renderQuickSuggestions() {
    const elLeaders = document.getElementById("quick-leaders-list");
    const elLocations = document.getElementById("quick-locations-list");
    const elStaff = document.getElementById("quick-staff-list");

    if (elLeaders) {
        elLeaders.innerHTML = QUICK_SUGGESTIONS.leaders.map(l => 
            `<span class="chip" onclick="navigator.clipboard.writeText('${escapeHtml(l)}'); showToast('Đã sao chép: ${escapeHtml(l)}', 'info')" title="Nhấp để sao chép">${escapeHtml(l)}</span>`
        ).join("");
    }

    if (elLocations) {
        elLocations.innerHTML = QUICK_SUGGESTIONS.locations.map(loc => 
            `<span class="chip" onclick="navigator.clipboard.writeText('${escapeHtml(loc)}'); showToast('Đã sao chép: ${escapeHtml(loc)}', 'info')" title="Nhấp để sao chép">${escapeHtml(loc)}</span>`
        ).join("");
    }

    if (elStaff) {
        elStaff.innerHTML = QUICK_SUGGESTIONS.staff.map(s => 
            `<span class="chip" onclick="navigator.clipboard.writeText('${escapeHtml(s)}'); showToast('Đã sao chép: ${escapeHtml(s)}', 'info')" title="Nhấp để sao chép">${escapeHtml(s)}</span>`
        ).join("");
    }
}

// Khởi chạy ứng dụng khi DOM tải xong
document.addEventListener("DOMContentLoaded", () => {
    renderQuickSuggestions();
    // Cập nhật số lượng lịch cũ trong kho
    fetch("/api/history")
        .then(r => r.json())
        .then(list => {
            const badge = document.getElementById("history-badge-count");
            if (badge && Array.isArray(list)) badge.textContent = list.length;
        })
        .catch(() => {});
    // Mặc định khởi tạo Tuần hiện tại (offset = 0) theo đúng yêu cầu người dùng
    switchWeekMode('current');
});
