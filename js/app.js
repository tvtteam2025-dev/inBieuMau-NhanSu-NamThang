document.addEventListener('DOMContentLoaded', () => {
    // Các phần tử giao diện
    const templateSelect = document.getElementById('template-select');
    const templateControl =
        document.getElementById('template-control') ||
        templateSelect.closest('.control-group');
    const recordIdInput = document.getElementById('record-id');
    const dataPreview = document.getElementById('data-preview');
    const documentContent = document.getElementById('document-content');

    // Nút bấm
    const btnReload = document.getElementById('btn-reload');
    const btnWord = document.getElementById('btn-word');

    // Login Elements
    const loginModal = document.getElementById('login-modal');
    const loginUsernameInput = document.getElementById('login-username');
    const loginPasswordInput = document.getElementById('login-password');
    const btnLogin = document.getElementById('btn-login');
    const loginError = document.getElementById('login-error');

    // Trạng thái hiện tại
    let currentData = null;
    let currentId = null;
    let currentHopDongId = null;

    // Các template dùng chung sẽ hiển thị cho mọi công ty.
    // Khi thêm mẫu chung, value phải trùng tên file trong thư mục templates,
    // nhưng không bao gồm phần mở rộng .docx.
    const COMMON_TEMPLATES = [
        // Ví dụ:
        {
            value: 'COMMON-camketbaomatthongtin',
            label: 'Cam kết bảo mật thông tin'
        }
    ];

    // Mỗi đơn vị chỉ nhìn thấy template dùng chung và template riêng của đơn vị đó.
    const TNM_TEMPLATES = [
        {
            value: 'TNM-danhgiathuviec',
            label: '1. TNM - Đánh giá thử việc'
        }
    ];

    const COMPANY_TEMPLATES = {
        'CÔNG TY TNHH ĐI XANH': [
            {
                value: 'DX-danhgiathuviec',
                label: '1. Đi Xanh - Đánh giá thử việc'
            }
        ],

        'CÔNG TY TNHH TM DV TRƯỜNG NHẬT MINH': TNM_TEMPLATES,

        // Giữ alias này để không lỗi nếu dữ liệu cũ trong Sheet đang ghi "NHẤT".
        'CÔNG TY TNHH TM DV TRƯỜNG NHẤT MINH': TNM_TEMPLATES
    };

    function normalizeCompanyName(value) {
        return String(value || '')
            .normalize('NFC')
            .replace(/\s+/g, ' ')
            .trim()
            .toUpperCase();
    }

    function setupTemplatesForCompany(data, requestedTemplate) {
        const companyName = normalizeCompanyName(
            data.congTyDonVi || data.main_congTyDonVi
        );
        const companyTemplates = COMPANY_TEMPLATES[companyName] || [];
        const templates = COMMON_TEMPLATES.concat(companyTemplates);

        templateSelect.innerHTML = '';

        if (templates.length === 0) {
            templateControl.style.display = 'none';
            btnWord.disabled = true;
            return null;
        }

        templates.forEach(function(item) {
            const option = document.createElement('option');
            option.value = item.value;
            option.textContent = item.label;
            templateSelect.appendChild(option);
        });

        const requestedIsAllowed = templates.some(function(item) {
            return item.value === requestedTemplate;
        });

        const selectedTemplate = requestedIsAllowed
            ? requestedTemplate
            : templates[0].value;

        templateSelect.value = selectedTemplate;
        templateControl.style.display = 'block';
        btnWord.disabled = false;

        return selectedTemplate;
    }

    function showMissingCompanyTemplate(data) {
        const companyName = String(
            data.congTyDonVi || data.main_congTyDonVi || 'không xác định'
        ).trim();
        const message = 'Chưa cấu hình biểu mẫu cho đơn vị: ' + companyName;

        Utils.showStatus(message, 'error');
        documentContent.textContent = message;
    }

    // Check Auth
    function checkAuth() {
        const params = Utils.getQueryParams();
        let user = sessionStorage.getItem('auth_user');
        let pass = sessionStorage.getItem('auth_pass');
        let hasAuthInUrl = false;

        // Nếu có trên URL, ưu tiên URL và lưu vào session
        if (params.username) {
            user = params.username;
            sessionStorage.setItem('auth_user', user);
            hasAuthInUrl = true;
        }
        if (params.password) {
            pass = params.password;
            sessionStorage.setItem('auth_pass', pass);
            hasAuthInUrl = true;
        }

        // Chỉ đọc thông tin đăng nhập từ URL một lần. Sau đó xóa khỏi thanh địa chỉ
        // để không ghi đè thông tin người dùng nhập lại và hạn chế lộ mật khẩu.
        if (hasAuthInUrl) {
            const cleanUrl = new URL(window.location.href);
            cleanUrl.searchParams.delete('username');
            cleanUrl.searchParams.delete('password');
            window.history.replaceState({}, '', cleanUrl);
        }

        return { user, pass };
    }

    function showLogin(errorMsg = '') {
        loginModal.style.display = 'flex';
        if (errorMsg) {
            loginError.textContent = errorMsg;
            loginError.style.display = 'block';
        } else {
            loginError.style.display = 'none';
        }
    }

    // Khởi tạo
    async function init() {
        // Kiểm tra Auth trước
        const auth = checkAuth();
        if (!auth.user || !auth.pass) {
            showLogin();
            return;
        }

        const params = Utils.getQueryParams();
        const template = params.template || templateSelect.value;
        const id = params.id;
        const hopDongId = params.hopDongId;

        if (!id) {
            Utils.showStatus('Vui lòng truyền id trên URL (VD: ?id=123)', 'error');
            recordIdInput.value = '';
            documentContent.innerHTML = 'Vui lòng cung cấp tham số URL hợp lệ (ví dụ: ?template=hop_dong_hop_tac&id=123)';
            return;
        }

        recordIdInput.value = id;
        await loadData(id, hopDongId, template);
    }

    async function renderTemplate(templateName) {
        try {
            documentContent.innerHTML = 'Đang tải tệp Word gốc...';
            await TemplateEngine.renderDocx(templateName, currentData, documentContent);
        } catch (error) {
            Utils.showStatus(error.message, 'error');
            documentContent.innerHTML = `<div class="status-message error" style="display:block;">${error.message}</div>`;
        }
    }

    // Gọi API và render HTML
    async function loadData(id, hopDongId, template) {
        // Tối ưu: Nếu đã có dữ liệu của chính ID này rồi, chỉ cần render lại template
        if (currentData && currentId === id && currentHopDongId === hopDongId) {
            const selectedTemplate = setupTemplatesForCompany(currentData, template);
            if (!selectedTemplate) {
                showMissingCompanyTemplate(currentData);
                return;
            }

            Utils.updateUrlParam('template', selectedTemplate);
            await renderTemplate(selectedTemplate);
            return;
        }

        Utils.hideStatus();
        documentContent.innerHTML = 'Đang tải dữ liệu...';
        dataPreview.textContent = 'Đang gọi API...';

        try {
            const auth = checkAuth();

            // Lấy data từ Apps Script
            const response = await API.fetchData(
                id,
                hopDongId,
                template,
                auth.user,
                auth.pass
            );

            if (!response.success) {
                if (response.auth_failed) {
                    sessionStorage.removeItem('auth_user');
                    sessionStorage.removeItem('auth_pass');
                    showLogin(response.message);
                    return;
                }

                const errorMessage = response.message || 'Lỗi API không xác định';
                Utils.showStatus(errorMessage, 'error');
                dataPreview.textContent = JSON.stringify(response, null, 2);
                documentContent.innerHTML = '';
                const errorElement = document.createElement('div');
                errorElement.className = 'status-message error';
                errorElement.style.display = 'block';
                errorElement.textContent = errorMessage;
                documentContent.appendChild(errorElement);
                return;
            }

            if (response.warnings && response.warnings.length > 0) {
                Utils.showStatus('Cảnh báo: ' + response.warnings.join(' ; '), 'warning');
            } else {
                Utils.showStatus('Tải dữ liệu thành công!', 'success');
                setTimeout(Utils.hideStatus, 3000);
            }

            currentData = response.data;
            currentId = id;
            currentHopDongId = hopDongId;
            dataPreview.textContent = JSON.stringify(response.data, null, 2);

            // Chọn đúng danh sách template theo công ty rồi mới render.
            const selectedTemplate = setupTemplatesForCompany(currentData, template);
            if (!selectedTemplate) {
                showMissingCompanyTemplate(currentData);
                return;
            }

            Utils.updateUrlParam('template', selectedTemplate);
            await renderTemplate(selectedTemplate);

        } catch (error) {
            Utils.showStatus(error.message, 'error');
            documentContent.innerHTML = 'Hệ thống gọi API thất bại: ' + error.message;
        }
    }

    // Lắng nghe sự kiện
    btnReload.addEventListener('click', () => {
        const id = recordIdInput.value;
        const template = templateSelect.value;
        const hopDongId = Utils.getQueryParams().hopDongId;
        if (id) {
            currentData = null; // Ép tải lại từ đầu
            currentId = null;
            currentHopDongId = null;
            loadData(id, hopDongId, template);
        }
    });

    templateSelect.addEventListener('change', (e) => {
        const newTemplate = e.target.value;
        const id = recordIdInput.value;
        const hopDongId = Utils.getQueryParams().hopDongId;
        Utils.updateUrlParam('template', newTemplate);
        if (id) {
            loadData(id, hopDongId, newTemplate);
        }
    });


    btnWord.addEventListener('click', () => {
        const id = recordIdInput.value;
        const template = templateSelect.value;
        ExportWord.generate(`${template}_${id}`);
    });




    btnLogin.addEventListener('click', () => {
        const u = loginUsernameInput.value.trim();
        const p = loginPasswordInput.value.trim();
        if (!u || !p) {
            loginError.textContent = 'Vui lòng nhập đầy đủ tên và mật khẩu';
            loginError.style.display = 'block';
            return;
        }
        sessionStorage.setItem('auth_user', u);
        sessionStorage.setItem('auth_pass', p);
        loginModal.style.display = 'none';
        init();
    });

    // Chạy app
    init();
});
