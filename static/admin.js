document.addEventListener('DOMContentLoaded', () => {
    const tableBody = document.getElementById('accounts-table-body');
    const accountTotalText = document.getElementById('account-total-text');
    const searchInput = document.getElementById('admin-search-input');
    const adminInput = document.getElementById('admin-account-str-input');
    const btnAdd = document.getElementById('btn-add-account');
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toast-message');

    let allAccounts = [];

    // 1. Kiểm tra sức khỏe API Server
    fetch('/health')
        .then(res => res.json())
        .then(data => {
            const serverStatus = document.getElementById('server-status');
            if (data.status === 'ok') {
                serverStatus.innerHTML = '<span class="pulse-dot"></span> SQLite Active';
                fetchAccounts();
            }
        })
        .catch(() => {
            const serverStatus = document.getElementById('server-status');
            serverStatus.style.background = 'rgba(239, 68, 68, 0.1)';
            serverStatus.style.color = '#f87171';
            serverStatus.style.borderColor = 'rgba(248, 113, 113, 0.2)';
            serverStatus.innerHTML = '⚠️ Server Mất Kết Nối';
        });

    // 2. Fetch danh sách tài khoản từ SQLite DB
    async function fetchAccounts() {
        try {
            const res = await fetch('/api/accounts');
            const data = await res.json();
            
            // Giải bóc mảng an toàn per User Rule #4
            allAccounts = Array.isArray(data) ? data : (data?.accounts || data?.items || []);
            renderTable(allAccounts);
        } catch (err) {
            showToast(`Lỗi khi nạp dữ liệu SQLite: ${err.message}`, 'error');
            tableBody.innerHTML = `<tr><td colspan="5" style="color: #ef4444; text-align: center; padding: 20px;">Lỗi nạp dữ liệu: ${err.message}</td></tr>`;
        }
    }

    // 3. Render Bảng Tài khoản
    function renderTable(accounts) {
        accountTotalText.textContent = `Tổng số: ${accounts.length} tài khoản`;

        if (!accounts || accounts.length === 0) {
            tableBody.innerHTML = `
                <tr>
                    <td colspan="5" style="text-align: center; padding: 30px; color: var(--text-muted);">
                        <i class="fa-regular fa-folder-open" style="font-size: 28px; margin-bottom: 8px;"></i>
                        <p>Chưa có tài khoản nào trong SQLite Database.</p>
                    </td>
                </tr>
            `;
            return;
        }

        tableBody.innerHTML = accounts.map((acc, idx) => {
            const maskedStr = escapeHtml(acc.account_str || '');
            const emailKey = escapeHtml(acc.email || '');

            return `
                <tr>
                    <td>${idx + 1}</td>
                    <td><span class="email-key"><i class="fa-regular fa-envelope"></i> ${emailKey}</span></td>
                    <td>
                        <span class="masked-str" title="${maskedStr}">${maskedStr}</span>
                    </td>
                    <td><i class="fa-regular fa-clock"></i> ${escapeHtml(acc.updated_at || '---')}</td>
                    <td>
                        <div class="btn-action-group">
                            <button class="btn btn-copy btn-sm btn-copy-str" data-str="${maskedStr}" title="Copy chuỗi tài khoản">
                                <i class="fa-regular fa-copy"></i> Copy
                            </button>
                            <button class="btn btn-danger btn-sm btn-delete-acc" data-email="${emailKey}" title="Xóa khỏi SQLite DB">
                                <i class="fa-solid fa-trash-can"></i> Xóa
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');

        // Gắn sự kiện Copy
        document.querySelectorAll('.btn-copy-str').forEach(btn => {
            btn.addEventListener('click', () => {
                const str = btn.getAttribute('data-str');
                navigator.clipboard.writeText(str).then(() => {
                    showToast('Đã sao chép chuỗi tài khoản!');
                });
            });
        });

        // Gắn sự kiện Xóa
        document.querySelectorAll('.btn-delete-acc').forEach(btn => {
            btn.addEventListener('click', async () => {
                const email = btn.getAttribute('data-email');
                if (!confirm(`Bạn có chắc chắn muốn xóa tài khoản '${email}' khỏi SQLite Database?`)) {
                    return;
                }

                try {
                    const res = await fetch(`/api/accounts/${encodeURIComponent(email)}`, {
                        method: 'DELETE'
                    });
                    const data = await res.json();
                    if (res.ok) {
                        showToast(`Đã xóa tài khoản ${email}`);
                        fetchAccounts();
                    } else {
                        throw new Error(data.detail || 'Lỗi khi xóa tài khoản');
                    }
                } catch (err) {
                    showToast(`Lỗi: ${err.message}`, 'error');
                }
            });
        });
    }

    // 4. Thêm tài khoản mới vào SQLite DB
    btnAdd.addEventListener('click', async () => {
        const str = adminInput.value.trim();
        if (!str) {
            showToast('Vui lòng nhập chuỗi tài khoản!', 'error');
            adminInput.focus();
            return;
        }

        try {
            const res = await fetch('/api/accounts/add', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ account_str: str })
            });

            const data = await res.json();
            if (res.ok) {
                showToast(`Đã lưu tài khoản '${data.email}' vào SQLite Database!`);
                adminInput.value = '';
                fetchAccounts();
            } else {
                throw new Error(data.detail || 'Lỗi lưu tài khoản');
            }
        } catch (err) {
            showToast(`Lỗi: ${err.message}`, 'error');
        }
    });

    // 5. Tìm kiếm realtime
    searchInput.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase().trim();
        if (!term) {
            renderTable(allAccounts);
            return;
        }

        const filtered = allAccounts.filter(acc => 
            (acc.email && acc.email.toLowerCase().includes(term)) ||
            (acc.account_str && acc.account_str.toLowerCase().includes(term))
        );
        renderTable(filtered);
    });

    function escapeHtml(str) {
        return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    function showToast(msg, type = 'success') {
        toastMessage.textContent = msg;
        toast.className = 'toast';
        if (type === 'error') {
            toast.style.background = '#ef4444';
        } else {
            toast.style.background = '#10b981';
        }
        toast.classList.remove('hidden');

        setTimeout(() => {
            toast.classList.add('hidden');
        }, 3500);
    }
});
