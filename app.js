// app.js - Logic cho Dashboard Báo Cáo Cận Date
const FILTER_STATE = {
    kho: new Set(),
    nhomhang: new Set(),
    trangthai: new Set()
};

let rawData = [];

// Khởi tạo
document.addEventListener('DOMContentLoaded', () => {
    if (typeof DATA_CANDATE !== 'undefined') {
        rawData = DATA_CANDATE;
    }
    
    // Thu thập danh sách giá trị filter duy nhất
    const khoSet = new Set();
    const nhomhangSet = new Set();
    const trangthaiSet = new Set();
    
    rawData.forEach(item => {
        if (item["MÃ KHO"]) khoSet.add(item["MÃ KHO"]);
        if (item["NHÓM HÀNG"]) nhomhangSet.add(item["NHÓM HÀNG"]);
        if (item["TRẠNG THÁI HSD"]) trangthaiSet.add(item["TRẠNG THÁI HSD"]);
    });
    
    // Khởi tạo các menu dropdown
    initDropdown('kho', Array.from(khoSet).sort(), true);
    initDropdown('nhomhang', Array.from(nhomhangSet).sort(), true);
    initDropdown('trangthai', Array.from(trangthaiSet).sort(), false, ["EXPIRED", "NEARLY EXPIRED"]); // Mặc định chỉ check Cận date và Hết hạn
    
    
    // Đóng dropdown khi click ra ngoài
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.filter-group')) {
            document.querySelectorAll('.filter-group').forEach(group => {
                group.classList.remove('active-dropdown');
            });
        }
    });

    document.getElementById('btnResetFilters').addEventListener('click', () => {
        selectAll('kho');
        selectAll('nhomhang');
        selectAll('trangthai');
    });

    document.getElementById('btnExportExcel').addEventListener('click', exportToExcel);
    if (typeof DATA_UPDATED_TIME !== 'undefined') {
        document.getElementById('last-update').innerText = `Cập nhật: ${DATA_UPDATED_TIME}`;
    } else {
        const now = new Date();
        document.getElementById('last-update').innerText = `Cập nhật: ${now.toLocaleTimeString('vi-VN')} ${now.toLocaleDateString('vi-VN')} (Local)`;
    }

    applyFiltersAndRender();
});

function initDropdown(filterType, values, defaultAll = true, defaultValues = []) {
    const listEl = document.getElementById(`list-${filterType}`);
    listEl.innerHTML = '';
    
    values.forEach(val => {
        if (!val) return;
        
        let isChecked = false;
        if (defaultAll) {
            isChecked = true;
        } else {
            isChecked = defaultValues.includes(val);
        }

        if (isChecked) {
            FILTER_STATE[filterType].add(val);
        }
        
        const label = document.createElement('label');
        label.className = 'dropdown-item';
        label.innerHTML = `
            <input type="checkbox" value="${val}" ${isChecked ? 'checked' : ''} onchange="handleFilterChange('${filterType}', this)">
            <span>${val}</span>
        `;
        listEl.appendChild(label);
    });
    updateFilterBadge(filterType);
}

function toggleDropdown(filterType) {
    const group = document.getElementById(`group-${filterType}`);
    const wasActive = group.classList.contains('active-dropdown');
    
    document.querySelectorAll('.filter-group').forEach(g => g.classList.remove('active-dropdown'));
    
    if (!wasActive) {
        group.classList.add('active-dropdown');
    }
}

function filterDropdownList(filterType, text) {
    const term = text.toLowerCase();
    const items = document.querySelectorAll(`#list-${filterType} .dropdown-item`);
    items.forEach(item => {
        const val = item.querySelector('span').innerText.toLowerCase();
        item.style.display = val.includes(term) ? 'flex' : 'none';
    });
}

function handleFilterChange(filterType, checkbox) {
    if (checkbox.checked) {
        FILTER_STATE[filterType].add(checkbox.value);
    } else {
        FILTER_STATE[filterType].delete(checkbox.value);
    }
    updateFilterBadge(filterType);
    applyFiltersAndRender();
}

function selectAll(filterType) {
    const inputs = document.querySelectorAll(`#list-${filterType} input[type="checkbox"]`);
    inputs.forEach(input => {
        input.checked = true;
        FILTER_STATE[filterType].add(input.value);
    });
    updateFilterBadge(filterType);
    applyFiltersAndRender();
}

function clearAll(filterType) {
    const inputs = document.querySelectorAll(`#list-${filterType} input[type="checkbox"]`);
    inputs.forEach(input => {
        input.checked = false;
        FILTER_STATE[filterType].delete(input.value);
    });
    updateFilterBadge(filterType);
    applyFiltersAndRender();
}

function updateFilterBadge(filterType) {
    const count = FILTER_STATE[filterType].size;
    const badge = document.getElementById(`badge-${filterType}`);
    badge.innerText = count;
    
    const triggerText = document.querySelector(`#btn-${filterType} .trigger-text`);
    if (count === 0) {
        triggerText.innerText = "Chưa chọn";
        triggerText.style.color = "#dc2626";
    } else if (count === document.querySelectorAll(`#list-${filterType} input`).length) {
        triggerText.innerText = "Tất cả";
        triggerText.style.color = "inherit";
    } else {
        triggerText.innerText = `Đã chọn (${count})`;
        triggerText.style.color = "var(--accent-blue)";
    }
}

function applyFiltersAndRender() {
    // 1. Lọc dữ liệu
    const filteredData = rawData.filter(item => {
        const matchKho = !item["MÃ KHO"] || FILTER_STATE.kho.has(item["MÃ KHO"]);
        const matchNhom = !item["NHÓM HÀNG"] || FILTER_STATE.nhomhang.has(item["NHÓM HÀNG"]);
        const matchTrangThai = !item["TRẠNG THÁI HSD"] || FILTER_STATE.trangthai.has(item["TRẠNG THÁI HSD"]);
        return matchKho && matchNhom && matchTrangThai;
    });



    // 3. Gom nhóm theo trạng thái
    const grouped = {
        "EXPIRED": [],
        "NEARLY EXPIRED": [],
        "USABLE": []
    };
    
    filteredData.forEach(item => {
        const status = item["TRẠNG THÁI HSD"] || "USABLE";
        if (grouped[status]) {
            grouped[status].push(item);
        } else {
            grouped["USABLE"].push(item); // Fallback
        }
    });

    // 4. Render Table (Hiển thị Hết hạn trước, rồi đến gần hết hạn, rồi mới đến usable)
    const tbody = document.getElementById('tbody-candate');
    tbody.innerHTML = '';
    
    // Hàm render 1 mảng
    const renderRows = (arr) => {
        arr.forEach(item => {
            const tr = document.createElement('tr');
            
            // Trang thai badge
            let statusBadge = '';
            if (item["TRẠNG THÁI HSD"] === "EXPIRED") {
                statusBadge = '<span style="color:white; background:#dc2626; padding:2px 6px; border-radius:4px; font-size:11px;">HẾT HẠN</span>';
            } else if (item["TRẠNG THÁI HSD"] === "NEARLY EXPIRED") {
                statusBadge = '<span style="color:white; background:#f59e0b; padding:2px 6px; border-radius:4px; font-size:11px;">GẦN HẾT HẠN</span>';
            } else {
                statusBadge = '<span style="color:white; background:#10b981; padding:2px 6px; border-radius:4px; font-size:11px;">CÓ THỂ XUẤT</span>';
            }
            
            tr.innerHTML = `
                <td>${item["MÃ KHO"] || ''}</td>
                <td>${item["MÃ HÀNG"] || ''}</td>
                <td style="max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${item["TÊN HÀNG"] || ''}">${item["TÊN HÀNG"] || ''}</td>
                <td>${item["NHÓM HÀNG"] || ''}</td>
                <td>${item["SỐ LÔ"] || ''}</td>
                <td>${item["NSX"] || ''}</td>
                <td>${item["HSD"] || ''}</td>
                <td class="text-right" style="font-weight:bold; ${item["SỐ NGÀY CÒN LẠI"] < 30 ? 'color:#dc2626;' : ''}">${item["SỐ NGÀY CÒN LẠI"]}</td>
                <td class="text-right" style="font-weight:bold;">${formatNum(item["SỐ LƯỢNG PL"])}</td>
                <td>${statusBadge}</td>
            `;
            tbody.appendChild(tr);
        });
    };

    renderRows(grouped["EXPIRED"]);
    renderRows(grouped["NEARLY EXPIRED"]);
    renderRows(grouped["USABLE"]);
}

function formatNum(num) {
    if (!num) return '0';
    return Number(num).toLocaleString('vi-VN', { maximumFractionDigits: 2 });
}

let sortAsc = true;
function sortTable(tableId, colIndex) {
    const table = document.getElementById(tableId);
    const tbody = table.tBodies[0];
    const rows = Array.from(tbody.querySelectorAll("tr"));
    
    sortAsc = !sortAsc;
    
    rows.sort((a, b) => {
        const cellA = a.cells[colIndex].innerText.trim();
        const cellB = b.cells[colIndex].innerText.trim();
        
        // Cố gắng parse số
        const numA = parseFloat(cellA.replace(/\./g, '').replace(/,/g, '.'));
        const numB = parseFloat(cellB.replace(/\./g, '').replace(/,/g, '.'));
        
        if (!isNaN(numA) && !isNaN(numB)) {
            return sortAsc ? numA - numB : numB - numA;
        }
        
        // So sánh chuỗi
        return sortAsc ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
    });
    
    // Render lại
    tbody.innerHTML = '';
    rows.forEach(row => tbody.appendChild(row));
}

function exportToExcel() {
    // 1. Lấy dữ liệu đã được lọc (dùng logic giống hàm applyFiltersAndRender)
    const currentFilteredData = rawData.filter(item => {
        const matchKho = !item["MÃ KHO"] || FILTER_STATE.kho.has(item["MÃ KHO"]);
        const matchNhom = !item["NHÓM HÀNG"] || FILTER_STATE.nhomhang.has(item["NHÓM HÀNG"]);
        const matchTrangThai = !item["TRẠNG THÁI HSD"] || FILTER_STATE.trangthai.has(item["TRẠNG THÁI HSD"]);
        return matchKho && matchNhom && matchTrangThai;
    });

    if (currentFilteredData.length === 0) {
        alert("Không có dữ liệu để xuất!");
        return;
    }

    // 2. Chuyển đổi dữ liệu thành Sheet
    const worksheet = XLSX.utils.json_to_sheet(currentFilteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "BaoCaoCanDate");

    // 3. Tải file xuống
    const now = new Date();
    const dateStr = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}`;
    XLSX.writeFile(workbook, `Bao_Cao_Can_Date_${dateStr}.xlsx`);
}
