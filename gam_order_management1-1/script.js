document.addEventListener('DOMContentLoaded', () => {
    // DOM 요소 가져오기
    const orderTableBody = document.getElementById('order-table-body');
    const newOrderBtn = document.getElementById('new-order-btn');
    const orderModal = document.getElementById('order-form-modal');
    const closeOrderModalBtn = orderModal.querySelector('.close-btn');
    const orderForm = document.getElementById('order-form');
    const formTitle = document.getElementById('form-title');
    const searchInput = document.getElementById('search-input');
    const dashboard = document.getElementById('dashboard');
    const resetViewBtn = document.getElementById('reset-view-btn');
    const toggleViewBtn = document.getElementById('toggle-view-btn');
    const orderTable = document.querySelector('.order-list table');

    // 가격 설정 관련 DOM 요소
    const priceSettingsBtn = document.getElementById('price-settings-btn');
    const priceModal = document.getElementById('price-settings-modal');
    const closePriceModalBtn = priceModal.querySelector('.close-btn');
    const priceSettingsForm = document.getElementById('price-settings-form');
    const sweetPriceInput = document.getElementById('sweet-price');
    const daebongPriceInput = document.getElementById('daebong-price');

    // 데이터 집계 관련 DOM 요소
    const aggregationBtn = document.getElementById('aggregation-btn');
    const aggregationModal = document.getElementById('aggregation-modal');
    const closeAggregationModalBtn = aggregationModal.querySelector('.close-btn');
    const aggregationControls = document.querySelector('.aggregation-controls');
    const aggregationResults = document.getElementById('aggregation-results');

    // 파일 입출력 관련 DOM 요소
    const saveDataBtn = document.getElementById('save-data-btn');
    const loadDataInput = document.getElementById('load-data-input');
    const exportScheduledBtn = document.getElementById('export-scheduled-btn');

    // 대시보드 요소
    const sentSweetP = document.getElementById('sent-sweet-persimmon');
    const sentDaebongP = document.getElementById('sent-daebong-persimmon');
    const unsentOrdersCount = document.getElementById('unsent-orders-count');
    const unsentSweetP = document.getElementById('unsent-sweet-persimmon');
    const unsentDaebongP = document.getElementById('unsent-daebong-persimmon');
    const delayedRedCount = document.getElementById('delayed-red-count');
    const delayedRedSweet = document.getElementById('delayed-red-sweet');
    const delayedRedDaebong = document.getElementById('delayed-red-daebong');
    const redThresholdInput = document.getElementById('red-threshold');
    const redThresholdLabel = document.getElementById('red-threshold-label');

    // 상태 관리 (기본값으로 초기화)
    let orders = [];
    let redThreshold = 7;
    let viewMode = 'compact';
    let prices = { sweet: 25000, daebong: 20000 };
    let currentSort = { column: 'orderDate', direction: 'desc' };
    let activeFilter = 'all';

    // UI 새로고침
    const refreshUI = () => {
        renderTable(getFilteredAndSortedOrders());
        updateDashboard();
        updateActiveFilterUI();
        updateViewModeUI();
        redThresholdInput.value = redThreshold;
        redThresholdLabel.textContent = redThreshold;
    };

    const updateActiveFilterUI = () => {
        document.querySelectorAll('.clickable-summary').forEach(item => {
            if (item.dataset.filter === activeFilter) {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });
    };

    const updateViewModeUI = () => {
        orderTable.className = 'view-' + viewMode;
        toggleViewBtn.textContent = viewMode === 'compact' ? '상세 보기' : '간단히 보기';
    };

    // 대시보드 업데이트
    const updateDashboard = () => {
        let sentSweet = 0, sentDaebong = 0;
        let unsentCount = 0, unsentSweet = 0, unsentDaebong = 0;
        let redCount = 0, redSweet = 0, redDaebong = 0;
        
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        orders.forEach(order => {
            const sweet = parseInt(order.sweetPersimmon) || 0;
            const daebong = parseInt(order.daebongPersimmon) || 0;

            if (['발송완료', '문제 발생'].includes(order.status)) {
                if (order.status === '발송완료') {
                    sentSweet += sweet;
                    sentDaebong += daebong;
                }
            } else {
                unsentCount++;
                unsentSweet += sweet;
                unsentDaebong += daebong;

                const orderDate = new Date(order.orderDate);
                const diffTime = today - orderDate;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                
                if (diffDays > redThreshold) {
                    redCount++;
                    redSweet += sweet;
                    redDaebong += daebong;
                }
            }
        });

        sentSweetP.textContent = sentSweet;
        sentDaebongP.textContent = sentDaebong;
        unsentOrdersCount.textContent = unsentCount;
        unsentSweetP.textContent = unsentSweet;
        unsentDaebongP.textContent = unsentDaebong;

        delayedRedCount.textContent = redCount;
        delayedRedSweet.textContent = redSweet;
        delayedRedDaebong.textContent = redDaebong;
    };

    // 날짜 및 숫자 형식 변환
    const formatDateWithDay = (dateString) => {
        if (!dateString) return '';
        const days = ['일', '월', '화', '수', '목', '금', '토'];
        const date = new Date(dateString + 'T00:00:00');
        const dayName = days[date.getDay()];
        return `${dateString} (${dayName})`;
    };

    const formatCurrency = (number) => {
        return number.toLocaleString('ko-KR') + '원';
    };

    // 테이블 렌더링
    const renderTable = (ordersToRender) => {
        orderTableBody.innerHTML = '';
        if (ordersToRender.length === 0) {
            orderTableBody.innerHTML = '<tr><td colspan="17" style="text-align:center;">표시할 주문이 없습니다.</td></tr>';
            return;
        }

        const isCompact = viewMode === 'compact';

        ordersToRender.forEach((order, index) => {
            const row = document.createElement('tr');
            if (!['발송완료', '문제 발생'].includes(order.status)) {
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                const orderDate = new Date(order.orderDate);
                const diffTime = today - orderDate;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

                if (diffDays > redThreshold) {
                    row.classList.add('urgency-critical');
                }
            }

            const senderNameHTML = isCompact ? 
                `<td class="col-name clickable-name" data-phone="${order.senderPhone}" data-address="${order.senderAddress}">${order.senderName}</td>` : 
                `<td class="col-name">${order.senderName}</td>`;

            const receiverNameHTML = isCompact ? 
                `<td class="col-name clickable-name" data-phone="${order.receiverPhone}" data-address="${order.receiverAddress}">${order.receiverName}</td>` : 
                `<td class="col-name">${order.receiverName}</td>`;
            
            const totalPrice = (parseInt(order.sweetPersimmon) || 0) * prices.sweet + (parseInt(order.daebongPersimmon) || 0) * prices.daebong;

            row.innerHTML = `
                <td>${index + 1}</td>
                <td class="col-group-end">${formatDateWithDay(order.orderDate)}</td>
                ${senderNameHTML}
                <td class="col-detail">${order.senderPhone}</td>
                <td class="col-detail col-group-end">${order.senderAddress}</td>
                ${receiverNameHTML}
                <td class="col-detail">${order.receiverPhone}</td>
                <td class="col-detail col-group-end">${order.receiverAddress}</td>
                <td>${order.sweetPersimmon}</td>
                <td>${order.daebongPersimmon}</td>
                <td>${formatCurrency(totalPrice)}</td>
                <td class="payment-status-cell col-group-end" data-id="${order.id}">
                    <span class="payment-status ${order.paymentStatus ? 'paid' : 'unpaid'}">
                        ${order.paymentStatus ? 'O' : 'X'}
                    </span>
                </td>
                <td>
                    <select class="status-select" data-id="${order.id}">
                        <option value="미발송" ${order.status === '미발송' ? 'selected' : ''}>미발송</option>
                        <option value="발송예정" ${order.status === '발송예정' ? 'selected' : ''}>발송예정</option>
                        <option value="발송완료" ${order.status === '발송완료' ? 'selected' : ''}>발송완료</option>
                        <option value="문제 발생" ${order.status === '문제 발생' ? 'selected' : ''}>문제 발생</option>
                    </select>
                </td>
                <td class="col-group-end">
                    ${order.status === '발송완료' ? 
                        `<input type="date" class="shipping-date-input" data-id="${order.id}" value="${order.shippingDate || ''}">` : 
                        (formatDateWithDay(order.shippingDate) || '')
                    }
                </td>
                <td>${order.notes || ''}</td>
                <td class="col-group-end">${order.referrer || ''}</td>
                <td>
                    <button class="edit-btn" data-id="${order.id}">수정</button>
                    <button class="delete-btn" data-id="${order.id}">삭제</button>
                </td>
            `;
            orderTableBody.appendChild(row);
        });
        document.getElementById('row-count-display').textContent = `총 ${ordersToRender.length}건`;
    };

    // 데이터 필터링 및 정렬
    const getFilteredAndSortedOrders = () => {
        let processedOrders = [...orders];
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const searchTerm = searchInput.value.toLowerCase();
        if (searchTerm) {
            processedOrders = processedOrders.filter(order => 
                Object.values(order).some(value => 
                    String(value).toLowerCase().includes(searchTerm)
                )
            );
        }

        let filteredOrders = [...processedOrders];

        switch (activeFilter) {
            case 'sent_sweet':
                filteredOrders = filteredOrders.filter(o => o.status === '발송완료' && (parseInt(o.sweetPersimmon) || 0) > 0);
                break;
            case 'sent_daebong':
                filteredOrders = filteredOrders.filter(o => o.status === '발송완료' && (parseInt(o.daebongPersimmon) || 0) > 0);
                break;
            case 'unsent_total':
                filteredOrders = filteredOrders.filter(o => !['발송완료', '문제 발생'].includes(o.status));
                break;
            case 'unsent_sweet':
                filteredOrders = filteredOrders.filter(o => !['발송완료', '문제 발생'].includes(o.status) && (parseInt(o.sweetPersimmon) || 0) > 0);
                break;
            case 'unsent_daebong':
                filteredOrders = filteredOrders.filter(o => !['발송완료', '문제 발생'].includes(o.status) && (parseInt(o.daebongPersimmon) || 0) > 0);
                break;
            case 'delayed_red':
                filteredOrders = filteredOrders.filter(o => {
                    if (['발송완료', '문제 발생'].includes(o.status)) return false;
                    const orderDate = new Date(o.orderDate);
                    const diffTime = today - orderDate;
                    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    return diffDays > redThreshold;
                });
                break;
        }

        filteredOrders.forEach(order => {
            order.totalPrice = (parseInt(order.sweetPersimmon) || 0) * prices.sweet + (parseInt(order.daebongPersimmon) || 0) * prices.daebong;
        });

        filteredOrders.sort((a, b) => {
            const column = currentSort.column;
            let valA = a[column];
            let valB = b[column];

            switch (column) {
                case 'sweetPersimmon':
                case 'daebongPersimmon':
                case 'totalPrice':
                    valA = parseInt(valA) || 0;
                    valB = parseInt(valB) || 0;
                    break;
            }

            let comparison = 0;
            if (valA > valB) comparison = 1; else if (valA < valB) comparison = -1;
            return currentSort.direction === 'asc' ? comparison : -comparison;
        });

        return filteredOrders;
    };

    // 모달 관리
    const openOrderModal = (orderId = null) => {
        orderForm.reset();
        if (orderId) {
            const order = orders.find(o => o.id == orderId);
            if(order){
                formTitle.textContent = '주문 수정';
                document.getElementById('order-id').value = order.id;
                document.getElementById('sender-name').value = order.senderName;
                document.getElementById('sender-phone').value = order.senderPhone;
                document.getElementById('sender-address').value = order.senderAddress;
                document.getElementById('receiver-name').value = order.receiverName;
                document.getElementById('receiver-phone').value = order.receiverPhone;
                document.getElementById('receiver-address').value = order.receiverAddress;
                document.getElementById('sweet-persimmon').value = order.sweetPersimmon;
                document.getElementById('daebong-persimmon').value = order.daebongPersimmon;
                document.getElementById('order-date').value = order.orderDate;
                document.getElementById('notes').value = order.notes || '';
                document.getElementById('referrer').value = order.referrer || '';
                document.getElementById('payment-status').checked = order.paymentStatus || false;
            }
        } else {
            formTitle.textContent = '새 주문';
            document.getElementById('order-id').value = '';
            document.getElementById('sender-address').value = '전남 순천시 낙안면 심내길 36';
            document.getElementById('order-date').value = new Date().toISOString().split('T')[0];
        }
        orderModal.style.display = 'flex';
    };
    const closeOrderModal = () => { orderModal.style.display = 'none'; };

    const openPriceModal = () => {
        sweetPriceInput.value = prices.sweet;
        daebongPriceInput.value = prices.daebong;
        priceModal.style.display = 'flex';
    };
    const closePriceModal = () => { priceModal.style.display = 'none'; };

    const openAggregationModal = () => {
        renderAggregationResults();
        aggregationModal.style.display = 'flex';
    };
    const closeAggregationModal = () => { aggregationModal.style.display = 'none'; };

    // 데이터 집계
    const calculateAggregations = () => {
        const results = {
            paid: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            unpaid: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            shipped: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            unshipped: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
        };

        orders.forEach(order => {
            const sweetCount = parseInt(order.sweetPersimmon) || 0;
            const daebongCount = parseInt(order.daebongPersimmon) || 0;
            const sweetValue = sweetCount * prices.sweet;
            const daebongValue = daebongCount * prices.daebong;

            if (order.paymentStatus) {
                results.paid.sweet += sweetValue;
                results.paid.daebong += daebongValue;
                results.paid.sweetCount += sweetCount;
                results.paid.daebongCount += daebongCount;
            } else {
                results.unpaid.sweet += sweetValue;
                results.unpaid.daebong += daebongValue;
                results.unpaid.sweetCount += sweetCount;
                results.unpaid.daebongCount += daebongCount;
            }

            if (order.status === '발송완료') {
                results.shipped.sweet += sweetValue;
                results.shipped.daebong += daebongValue;
                results.shipped.sweetCount += sweetCount;
                results.shipped.daebongCount += daebongCount;
            } else {
                results.unshipped.sweet += sweetValue;
                results.unshipped.daebong += daebongValue;
                results.unshipped.sweetCount += sweetCount;
                results.unshipped.daebongCount += daebongCount;
            }
        });

        results.paid.total = results.paid.sweet + results.paid.daebong;
        results.unpaid.total = results.unpaid.sweet + results.unpaid.daebong;
        results.shipped.total = results.shipped.sweet + results.shipped.daebong;
        results.unshipped.total = results.unshipped.sweet + results.unshipped.daebong;

        return results;
    };

    const renderAggregationResults = () => {
        const results = calculateAggregations();
        const criteria = document.querySelector('input[name="agg-criteria"]:checked').value;

        let html = '';
        let data, cat1, cat2, cat1Label, cat2Label;

        if (criteria === 'payment') {
            data = results;
            cat1 = 'paid';
            cat2 = 'unpaid';
            cat1Label = '입금 완료';
            cat2Label = '미입금';
            html = `<h3>입금여부 기준 집계</h3>`;
        } else { // shipping
            data = results;
            cat1 = 'shipped';
            cat2 = 'unshipped';
            cat1Label = '발송 완료';
            cat2Label = '미발송';
            html = `<h3>발송여부 기준 집계</h3>`;
        }

        const totalSweetCount = data[cat1].sweetCount + data[cat2].sweetCount;
        const totalDaebongCount = data[cat1].daebongCount + data[cat2].daebongCount;
        const totalSweetValue = data[cat1].sweet + data[cat2].sweet;
        const totalDaebongValue = data[cat1].daebong + data[cat2].daebong;
        const grandTotal = data[cat1].total + data[cat2].total;

        html += `
            <table>
                <thead>
                    <tr>
                        <th>항목</th>
                        <th>단감(금액)</th>
                        <th>단감(수량)</th>
                        <th>대봉(금액)</th>
                        <th>대봉(수량)</th>
                        <th>합계</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>${cat1Label}</td>
                        <td>${formatCurrency(data[cat1].sweet)}</td>
                        <td>${data[cat1].sweetCount}</td>
                        <td>${formatCurrency(data[cat1].daebong)}</td>
                        <td>${data[cat1].daebongCount}</td>
                        <td>${formatCurrency(data[cat1].total)}</td>
                    </tr>
                    <tr>
                        <td>${cat2Label}</td>
                        <td>${formatCurrency(data[cat2].sweet)}</td>
                        <td>${data[cat2].sweetCount}</td>
                        <td>${formatCurrency(data[cat2].daebong)}</td>
                        <td>${data[cat2].daebongCount}</td>
                        <td>${formatCurrency(data[cat2].total)}</td>
                    </tr>
                </tbody>
                <tfoot>
                    <tr>
                        <th>총계</th>
                        <td>${formatCurrency(totalSweetValue)}</td>
                        <td>${totalSweetCount}</td>
                        <td>${formatCurrency(totalDaebongValue)}</td>
                        <td>${totalDaebongCount}</td>
                        <td>${formatCurrency(grandTotal)}</td>
                    </tr>
                </tfoot>
            </table>
        `;
        aggregationResults.innerHTML = html;
    };

    // 파일 저장/불러오기
    const saveStateToFile = () => {
        const appState = {
            orders: orders,
            prices: prices,
            redThreshold: redThreshold,
            viewMode: viewMode
        };
        const blob = new Blob([JSON.stringify(appState, null, 2)], { type: 'application/json' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'gam_data.json';
        link.click();
    };

    const loadStateFromFile = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const appState = JSON.parse(event.target.result);
                orders = appState.orders || [];
                prices = appState.prices || { sweet: 25000, daebong: 20000 };
                redThreshold = appState.redThreshold || 7;
                viewMode = appState.viewMode || 'compact';
                alert('데이터를 성공적으로 불러왔습니다.');
            } catch (error) {
                alert('파일을 읽는 중 오류가 발생했습니다. 유효한 데이터 파일인지 확인해주세요.');
                console.error("Failed to parse JSON file:", error);
            }
            refreshUI();
        };
        reader.readAsText(file);
        e.target.value = ''; // 동일한 파일을 다시 불러올 수 있도록 초기화
    };

    // 이벤트 리스너
    newOrderBtn.addEventListener('click', () => openOrderModal());
    closeOrderModalBtn.addEventListener('click', closeOrderModal);
    priceSettingsBtn.addEventListener('click', openPriceModal);
    closePriceModalBtn.addEventListener('click', closePriceModal);
    aggregationBtn.addEventListener('click', openAggregationModal);
    closeAggregationModalBtn.addEventListener('click', closeAggregationModal);
    aggregationControls.addEventListener('change', renderAggregationResults);
    saveDataBtn.addEventListener('click', saveStateToFile);
    loadDataInput.addEventListener('change', loadStateFromFile);

    window.addEventListener('click', (e) => { 
        if (e.target === orderModal) closeOrderModal();
        if (e.target === priceModal) closePriceModal();
        if (e.target === aggregationModal) closeAggregationModal();
    });

    toggleViewBtn.addEventListener('click', () => {
        viewMode = viewMode === 'compact' ? 'expanded' : 'compact';
        refreshUI();
        saveStateToLocalStorage();
    });

    resetViewBtn.addEventListener('click', () => {
        activeFilter = 'all';
        searchInput.value = '';
        refreshUI();
        saveStateToLocalStorage();
    });

    dashboard.addEventListener('click', (e) => {
        const target = e.target.closest('.clickable-summary');
        if (target) {
            const filter = target.dataset.filter;
            if (activeFilter === filter) {
                activeFilter = 'all';
            } else {
                activeFilter = filter;
            }
            refreshUI();
        }
    });

    orderForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const id = document.getElementById('order-id').value;
        const orderData = {
            senderName: document.getElementById('sender-name').value,
            senderPhone: document.getElementById('sender-phone').value,
            senderAddress: document.getElementById('sender-address').value,
            receiverName: document.getElementById('receiver-name').value,
            receiverPhone: document.getElementById('receiver-phone').value,
            receiverAddress: document.getElementById('receiver-address').value,
            sweetPersimmon: document.getElementById('sweet-persimmon').value,
            daebongPersimmon: document.getElementById('daebong-persimmon').value,
            orderDate: document.getElementById('order-date').value,
            notes: document.getElementById('notes').value,
            referrer: document.getElementById('referrer').value,
            paymentStatus: document.getElementById('payment-status').checked
        };
        if (id) {
            const index = orders.findIndex(o => o.id == id);
            if(index !== -1) orders[index] = { ...orders[index], ...orderData };
        } else {
            orders.push({ id: Date.now(), status: '미발송', shippingDate: null, ...orderData });
        }
        refreshUI();
        closeOrderModal();
        saveStateToLocalStorage();
    });

    priceSettingsForm.addEventListener('submit', (e) => {
        e.preventDefault();
        prices.sweet = parseInt(sweetPriceInput.value) || 0;
        prices.daebong = parseInt(daebongPriceInput.value) || 0;
        refreshUI();
        closePriceModal();
        saveStateToLocalStorage();
    });

    orderTableBody.addEventListener('click', (e) => {
        const targetCell = e.target.closest('td');
        if (!targetCell) return;

        const id = targetCell.dataset.id || e.target.dataset.id;

        if (targetCell.classList.contains('payment-status-cell')) {
            const index = orders.findIndex(o => o.id == id);
            if (index !== -1) {
                orders[index].paymentStatus = !orders[index].paymentStatus;
                refreshUI();
                saveStateToLocalStorage();
            }
        } else if (e.target.classList.contains('edit-btn')) {
            openOrderModal(parseInt(id));
        } else if (e.target.classList.contains('delete-btn')) {
            if (confirm('정말로 이 주문을 삭제하시겠습니까?')) {
                orders = orders.filter(o => o.id != id);
                refreshUI();
                saveStateToLocalStorage();
            }
        } else if (targetCell.classList.contains('clickable-name')) {
            const phone = targetCell.dataset.phone;
            const address = targetCell.dataset.address;
            alert(`연락처: ${phone}\n주소: ${address}`);
        }
    });

    orderTableBody.addEventListener('change', (e) => {
        const id = e.target.dataset.id;
        if (e.target.classList.contains('status-select')) {
            const index = orders.findIndex(o => o.id == id);
            if(index !== -1) {
                orders[index].status = e.target.value;
                if (e.target.value === '발송완료') {
                    if (!orders[index].shippingDate) {
                        orders[index].shippingDate = new Date().toISOString().split('T')[0];
                    }
                } else {
                    orders[index].shippingDate = null;
                }
                refreshUI();
                saveStateToLocalStorage();
            }
        } else if (e.target.classList.contains('shipping-date-input')) {
            const index = orders.findIndex(o => o.id == id);
            if(index !== -1) {
                orders[index].shippingDate = e.target.value;
                updateDashboard();
                saveStateToLocalStorage();
            }
        }
    });

    searchInput.addEventListener('input', () => renderTable(getFilteredAndSortedOrders()));

    document.querySelectorAll('th[data-sort]').forEach(header => {
        header.addEventListener('click', () => {
            const column = header.dataset.sort;
            if (currentSort.column === column) {
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort.column = column;
                currentSort.direction = 'desc';
            }
            refreshUI();
        });
    });

    redThresholdInput.addEventListener('change', () => {
        redThreshold = parseInt(redThresholdInput.value) || 7;
        refreshUI();
        saveStateToLocalStorage();
    });

    exportScheduledBtn.addEventListener('click', () => {
        const scheduledOrders = orders.filter(o => o.status === '발송예정');
        if (scheduledOrders.length === 0) {
            alert('발송예정인 주문이 없습니다.');
            return;
        }
        const dataToExport = scheduledOrders.map(order => ({
            '보내는 분': order.senderName,
            '보내는 분 연락처': order.senderPhone,
            '보내는 분 주소': order.senderAddress,
            '받는 분': order.receiverName,
            '받는 분 연락처': order.receiverPhone,
            '받는 분 주소': order.receiverAddress,
            '단감(개)': order.sweetPersimmon,
            '대봉(개)': order.daebongPersimmon
        }));
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '발송예정 목록');
        XLSX.writeFile(workbook, '발송예정_주문목록.xlsx');
    });
    
    // 로컬 스토리지에 상태 저장
    const saveStateToLocalStorage = () => {
        const appState = {
            orders: orders,
            prices: prices,
            redThreshold: redThreshold,
            viewMode: viewMode,
            currentSort: currentSort,
            activeFilter: activeFilter
        };
        localStorage.setItem('gamAppState', JSON.stringify(appState));
    };

    // 로컬 스토리지에서 상태 불러오기
    const loadStateFromLocalStorage = () => {
        const savedState = localStorage.getItem('gamAppState');
        if (savedState) {
            try {
                const appState = JSON.parse(savedState);
                orders = appState.orders || [];
                prices = appState.prices || { sweet: 25000, daebong: 20000 };
                redThreshold = appState.redThreshold || 7;
                viewMode = appState.viewMode || 'compact';
                currentSort = appState.currentSort || { column: 'orderDate', direction: 'desc' };
                activeFilter = appState.activeFilter || 'all';
            } catch (error) {
                console.error("Failed to parse app state from local storage:", error);
                // 오류 발생 시 기본값으로 초기화
                orders = [];
                redThreshold = 7;
                viewMode = 'compact';
                prices = { sweet: 25000, daebong: 20000 };
                currentSort = { column: 'orderDate', direction: 'desc' };
                activeFilter = 'all';
            }
        }
    };

    // 초기화
    const initialize = () => {
        loadStateFromLocalStorage(); // 로컬 스토리지에서 상태 로드
        refreshUI();
    };

    initialize();
    // 페이지를 닫거나 새로고침하기 전에 상태 저장
    window.addEventListener('beforeunload', saveStateToLocalStorage);
});