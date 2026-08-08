document.addEventListener('DOMContentLoaded', () => {
    const accountInput = document.getElementById('account-str-input');
    const keywordInput = document.getElementById('keyword-input');
    const limitInput = document.getElementById('limit-input');
    const mockToggle = document.getElementById('mock-toggle');
    const btnFetch = document.getElementById('btn-fetch');
    const btnText = btnFetch.querySelector('.btn-text');
    const btnLoader = btnFetch.querySelector('.btn-loader');

    const otpCodeDisplay = document.getElementById('otp-code-display');
    const btnCopyOtp = document.getElementById('btn-copy-otp');
    const otpTime = document.getElementById('otp-time');
    const otpSubjectText = document.getElementById('otp-subject-text');
    const otpSenderText = document.getElementById('otp-sender-text');

    const emailCountText = document.getElementById('email-count-text');
    const emailsListContainer = document.getElementById('emails-list-container');
    const filterEmailsInput = document.getElementById('filter-emails-input');
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toast-message');

    let allFetchedEmails = [];

    const savedAccountsSelect = document.getElementById('saved-accounts-select');

    // Nạp danh sách tài khoản đã lưu trong SQLite DB
    function loadSavedAccountsFromDB() {
        fetch('/api/accounts')
            .then(res => res.json())
            .then(data => {
                // Giải bóc mảng an toàn per User Rule #4
                const accounts = Array.isArray(data) ? data : (data?.accounts || data?.items || []);
                savedAccountsSelect.innerHTML = '<option value="">-- Chọn tài khoản đã lưu trong Database --</option>';
                accounts.forEach(acc => {
                    const opt = document.createElement('option');
                    opt.value = acc.account_str;
                    opt.textContent = `${acc.email} (Lần cuối: ${acc.updated_at || '---'})`;
                    savedAccountsSelect.appendChild(opt);
                });
            })
            .catch(err => console.error('Lỗi khi nạp tài khoản SQLite DB:', err));
    }

    if (savedAccountsSelect) {
        savedAccountsSelect.addEventListener('change', (e) => {
            if (e.target.value) {
                accountInput.value = e.target.value;
            }
        });
    }

    // 1. Kiểm tra sức khỏe API Server & Nạp DB
    fetch('/health')
        .then(res => res.json())
        .then(data => {
            const serverStatus = document.getElementById('server-status');
            if (data.status === 'ok') {
                serverStatus.innerHTML = '<span class="pulse-dot"></span> Server Sẵn sàng';
                loadSavedAccountsFromDB();
            }
        })
        .catch(() => {
            const serverStatus = document.getElementById('server-status');
            serverStatus.style.background = 'rgba(239, 68, 68, 0.1)';
            serverStatus.style.color = '#f87171';
            serverStatus.style.borderColor = 'rgba(248, 113, 113, 0.2)';
            serverStatus.innerHTML = '⚠️ Server Mất Kết Nối';
        });


    // 2. Xử lý nút Lấy mã OTP & Đọc Mail
    btnFetch.addEventListener('click', async () => {
        const accountStr = accountInput.value.trim();
        if (!accountStr) {
            showToast('Vui lòng nhập dòng thông tin tài khoản Outlook!', 'error');
            accountInput.focus();
            return;
        }

        // Toggle UI Loading
        btnFetch.disabled = true;
        btnText.classList.add('hidden');
        btnLoader.classList.remove('hidden');

        try {
            const payload = {
                account_str: accountStr,
                keyword: keywordInput.value.trim() || null,
                limit: parseInt(limitInput.value) || 15,
                use_mock: mockToggle.checked
            };

            const response = await fetch('/api/get-code', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.detail || 'Lỗi kết nối hòm thư Outlook');
            }

            // Render Hero OTP Card
            if (data.otp_code) {
                otpCodeDisplay.textContent = data.otp_code;
                otpTime.textContent = data.received_time || 'Vừa xong';
                otpSubjectText.textContent = data.subject || '---';
                otpSenderText.textContent = data.sender || '---';
                showToast(`Đã tìm thấy mã OTP: ${data.otp_code}`);
            } else {
                otpCodeDisplay.textContent = 'N/A';
                otpTime.textContent = '---';
                otpSubjectText.textContent = data.subject || 'Không thấy mã OTP trong email gần đây';
                otpSenderText.textContent = data.sender || '---';
                showToast('Đã đọc hòm thư nhưng không tìm thấy mã OTP nào.', 'info');
            }

            // Render Email Dashboard List
            allFetchedEmails = data.all_messages || [];
            renderEmailList(allFetchedEmails);
            loadSavedAccountsFromDB();


        } catch (err) {
            showToast(`Lỗi: ${err.message}`, 'error');
            otpCodeDisplay.textContent = 'ERROR';
            otpSubjectText.textContent = err.message;
        } finally {
            btnFetch.disabled = false;
            btnText.classList.remove('hidden');
            btnLoader.classList.add('hidden');
        }
    });

    // 3. Render danh sách Email
    function renderEmailList(emails) {
        emailCountText.textContent = `Hiển thị ${emails.length} email`;

        if (!emails || emails.length === 0) {
            emailsListContainer.innerHTML = `
                <div class="empty-state">
                    <i class="fa-regular fa-folder-open empty-icon"></i>
                    <p>Không có email nào trong hòm thư.</p>
                </div>
            `;
            return;
        }

        emailsListContainer.innerHTML = emails.map(email => {
            const hasOtp = email.otp_codes && email.otp_codes.length > 0;
            const otpTag = hasOtp 
                ? `<span class="tag-otp"><i class="fa-solid fa-key"></i> OTP: ${email.otp_codes.join(', ')}</span>`
                : '';
            const attTag = email.has_attachments 
                ? `<span><i class="fa-solid fa-paperclip"></i> Có tệp đính kèm</span>`
                : '';

            return `
                <div class="email-item" data-email-id="${email.id}">
                    <div class="email-item-header">
                        <span class="email-subject">${escapeHtml(email.subject || 'Không có tiêu đề')}</span>
                        ${otpTag}
                    </div>
                    <div class="email-sender"><i class="fa-regular fa-user"></i> ${escapeHtml(email.sender || '---')}</div>
                    <div class="email-preview">${escapeHtml(email.body_preview || '')}</div>
                    <div class="email-footer">
                        <span><i class="fa-regular fa-clock"></i> ${escapeHtml(email.created_date || '')}</span>
                        ${attTag}
                    </div>
                </div>
            `;
        }).join('');

        // Gắn sự kiện click vào từng email item để mở modal
        document.querySelectorAll('.email-item').forEach(item => {
            item.addEventListener('click', () => {
                const id = item.getAttribute('data-email-id');
                const targetEmail = emails.find(e => e.id === id);
                if (targetEmail) {
                    openEmailModal(targetEmail);
                }
            });
        });
    }

    // Modal elements
    const emailModal = document.getElementById('email-modal');
    const btnCloseModal = document.getElementById('btn-close-modal');
    const modalSubject = document.getElementById('modal-subject');
    const modalSender = document.getElementById('modal-sender');
    const modalDate = document.getElementById('modal-date');
    const modalOtpTag = document.getElementById('modal-otp-tag');
    const modalBodyContent = document.getElementById('modal-body-content');

    function openEmailModal(email) {
        modalSubject.textContent = email.subject || 'Không có tiêu đề';
        modalSender.textContent = email.sender || '---';
        modalDate.innerHTML = `<i class="fa-regular fa-clock"></i> ${email.created_date || ''}`;
        
        if (email.otp_codes && email.otp_codes.length > 0) {
            modalOtpTag.innerHTML = `<span class="tag-otp"><i class="fa-solid fa-key"></i> OTP: ${email.otp_codes.join(', ')}</span>`;
        } else {
            modalOtpTag.innerHTML = '';
        }

        modalBodyContent.textContent = email.body || email.body_preview || 'Không có nội dung.';
        emailModal.classList.remove('hidden');
    }

    btnCloseModal.addEventListener('click', () => {
        emailModal.classList.add('hidden');
    });

    emailModal.addEventListener('click', (e) => {
        if (e.target === emailModal) {
            emailModal.classList.add('hidden');
        }
    });


    // 4. Lọc email realtime trên UI
    filterEmailsInput.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase().trim();
        if (!term) {
            renderEmailList(allFetchedEmails);
            return;
        }
        const filtered = allFetchedEmails.filter(email => 
            (email.subject && email.subject.toLowerCase().includes(term)) ||
            (email.sender && email.sender.toLowerCase().includes(term)) ||
            (email.body_preview && email.body_preview.toLowerCase().includes(term))
        );
        renderEmailList(filtered);
    });

    // 5. Copy OTP 1-click
    btnCopyOtp.addEventListener('click', () => {
        const code = otpCodeDisplay.textContent.trim();
        if (!code || code === '------' || code === 'ERROR' || code === 'N/A') {
            showToast('Chưa có mã OTP hợp lệ để sao chép!', 'error');
            return;
        }
        navigator.clipboard.writeText(code).then(() => {
            showToast(`Đã sao chép mã OTP: ${code}`);
        });
    });

    // Helper: Escape HTML
    function escapeHtml(str) {
        return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    // Helper: Toast Notifications
    function showToast(msg, type = 'success') {
        toastMessage.textContent = msg;
        toast.className = 'toast';
        if (type === 'error') {
            toast.style.background = '#ef4444';
        } else if (type === 'info') {
            toast.style.background = '#3b82f6';
        } else {
            toast.style.background = '#10b981';
        }
        toast.classList.remove('hidden');

        setTimeout(() => {
            toast.classList.add('hidden');
        }, 3500);
    }
});
