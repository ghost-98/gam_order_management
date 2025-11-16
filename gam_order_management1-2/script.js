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

    // 배송 관리 모달 관련 DOM 요소
    const shippingModal = document.getElementById('shipping-management-modal');
    const closeShippingModalBtn = shippingModal.querySelector('.close-btn');
    const shippingModalSaveBtn = document.getElementById('shipping-modal-save-btn');
    const shippingOrderIdInput = document.getElementById('shipping-order-id');
    const shippingSweetTotal = document.getElementById('shipping-sweet-total');
    const shippingDaebongTotal = document.getElementById('shipping-daebong-total');
    const shippingSweetUnsent = document.getElementById('shipping-sweet-unsent');
    const shippingDaebongUnsent = document.getElementById('shipping-daebong-unsent');
    const shippingSweetScheduledList = document.getElementById('shipping-sweet-scheduled-list');
    const shippingDaebongScheduledList = document.getElementById('shipping-daebong-scheduled-list');
    const shippingSweetSentList = document.getElementById('shipping-sweet-sent-list');
    const shippingDaebongSentList = document.getElementById('shipping-daebong-sent-list');
    const shippingAddForms = document.querySelectorAll('.shipping-add-form');


    // 대시보드 요소
    const redThresholdInput = document.getElementById('red-threshold');
    const redThresholdLabel = document.getElementById('red-threshold-label-db');

    // 상태 관리 (기본값으로 초기화)
    let appState = {
        orders: [],
        redThreshold: 7,
        viewMode: 'compact',
        prices: { sweet: 25000, daebong: 20000 },
        currentSort: { column: 'orderDate', direction: 'desc' },
        currentFilter: null,
    };
    let tempShippingDetails = {}; // 배송관리 모달용 임시 데이터
    let editingShippingItem = null; // 배송관리 모달에서 수정중인 아이템 정보

    // 주문의 배송 상태(완료, 예정, 미발송)를 계산하는 헬퍼 함수
    const getOrderShippingStats = (order) => {
        const sweetTotal = parseInt(order.sweetPersimmon) || 0;
        const daebongTotal = parseInt(order.daebongPersimmon) || 0;
        const details = order.shippingDetails || { sweetPersimmon: [], daebongPersimmon: [] };

        let sentSweet = 0, scheduledSweet = 0;
        (details.sweetPersimmon || []).forEach(item => {
            if (item.status === '발송완료') sentSweet += item.count;
            else if (item.status === '발송예정') scheduledSweet += item.count;
        });

        let sentDaebong = 0, scheduledDaebong = 0;
        (details.daebongPersimmon || []).forEach(item => {
            if (item.status === '발송완료') sentDaebong += item.count;
            else if (item.status === '발송예정') scheduledDaebong += item.count;
        });

        const unsentSweet = sweetTotal - sentSweet - scheduledSweet;
        const unsentDaebong = daebongTotal - sentDaebong - scheduledDaebong;

        return {
            sentSweet, scheduledSweet, unsentSweet,
            sentDaebong, scheduledDaebong, unsentDaebong,
            totalSweet: sweetTotal,
            totalDaebong: daebongTotal
        };
    };

    // Helper function to check if an order is delayed
    const isOrderDelayed = (order, today) => {
        const stats = getOrderShippingStats(order);
        if (stats.unsentSweet <= 0 && stats.unsentDaebong <= 0) {
            return false; // Not delayed if everything is sent or scheduled
        }

        const orderDate = new Date(order.orderDate + 'T00:00:00');
        const diffTime = today.getTime() - orderDate.getTime();
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)); // Use floor for full days

        return diffDays >= appState.redThreshold;
    };

    // UI 새로고침
    const refreshUI = () => {
        renderTable(getFilteredAndSortedOrders());
        updateDashboard();
        updateViewModeUI();
        redThresholdInput.value = appState.redThreshold;
        if(redThresholdLabel) redThresholdLabel.textContent = appState.redThreshold;
    };

    const updateViewModeUI = () => {
        orderTable.className = 'view-' + appState.viewMode;
        toggleViewBtn.textContent = appState.viewMode === 'compact' ? '상세 보기' : '간단히 보기';
    };

    // 대시보드 업데이트
    const updateDashboard = () => {
        let totalOrders = 0;
        let totalSweetPersimmons = 0;
        let totalDaebongPersimmons = 0;
        let completedOrders = 0, totalSentSweet = 0, totalSentDaebong = 0;
        let unsentOrders = 0, totalUnsentSweet = 0, totalUnsentDaebong = 0;
        let scheduledOrders = 0, totalScheduledSweet = 0, totalScheduledDaebong = 0;
        let delayedOrders = 0, delayedUnsentSweet = 0, delayedUnsentDaebong = 0;
        
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        appState.orders.forEach(order => {
            totalOrders++;
            const stats = getOrderShippingStats(order);

            totalSweetPersimmons += stats.totalSweet;
            totalDaebongPersimmons += stats.totalDaebong;

            totalSentSweet += stats.sentSweet;
            totalSentDaebong += stats.sentDaebong;
            totalScheduledSweet += stats.scheduledSweet;
            totalScheduledDaebong += stats.scheduledDaebong;
            totalUnsentSweet += stats.unsentSweet;
            totalUnsentDaebong += stats.unsentDaebong;

            const totalBoxes = stats.totalSweet + stats.totalDaebong;
            const sentBoxes = stats.sentSweet + stats.sentDaebong;
            if (totalBoxes > 0 && totalBoxes === sentBoxes) {
                completedOrders++;
            }

            if (stats.unsentSweet > 0 || stats.unsentDaebong > 0) {
                unsentOrders++;
            }
            if (stats.scheduledSweet > 0 || stats.scheduledDaebong > 0) {
                scheduledOrders++;
            }
            
            if (isOrderDelayed(order, today)) {
                delayedOrders++;
                delayedUnsentSweet += stats.unsentSweet;
                delayedUnsentDaebong += stats.unsentDaebong;
            }
        });

        // 총 주문 그룹
        document.getElementById('db-total-orders').textContent = totalOrders;
        document.getElementById('db-total-sweet').textContent = totalSweetPersimmons;
        document.getElementById('db-total-daebong').textContent = totalDaebongPersimmons;
        
        // 주문 상태별 현황
        document.getElementById('db-completed-orders').textContent = completedOrders;
        document.getElementById('db-unsent-orders').textContent = unsentOrders;
        document.getElementById('db-scheduled-orders').textContent = scheduledOrders;

        // 상태별 박스 현황 그리드
        document.getElementById('db-total-sent-boxes').textContent = totalSentSweet + totalSentDaebong;
        document.getElementById('db-total-unsent-boxes').textContent = totalUnsentSweet + totalUnsentDaebong;
        document.getElementById('db-total-scheduled-boxes').textContent = totalScheduledSweet + totalScheduledDaebong;
        
        document.getElementById('db-sent-sweet').textContent = totalSentSweet;
        document.getElementById('db-unsent-sweet').textContent = totalUnsentSweet;
        document.getElementById('db-scheduled-sweet').textContent = totalScheduledSweet;

        document.getElementById('db-sent-daebong').textContent = totalSentDaebong;
        document.getElementById('db-unsent-daebong').textContent = totalUnsentDaebong;
        document.getElementById('db-scheduled-daebong').textContent = totalScheduledDaebong;
        
        // 지연 주문 현황
        document.getElementById('db-delayed-orders').textContent = delayedOrders;
        document.getElementById('db-delayed-sweet').textContent = delayedUnsentSweet;
        document.getElementById('db-delayed-daebong').textContent = delayedUnsentDaebong;
        
        if (redThresholdLabel) redThresholdLabel.textContent = appState.redThreshold;
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

    const createShippingStatusVisual = (order) => {
        const stats = getOrderShippingStats(order);
        const totalBoxes = stats.totalSweet + stats.totalDaebong;

        if (totalBoxes === 0) {
            return '<span>-</span>';
        }

        const sentCount = stats.sentSweet + stats.sentDaebong;
        const scheduledCount = stats.scheduledSweet + stats.scheduledDaebong;
        const unsentCount = totalBoxes - sentCount - scheduledCount;

        const unsentPercent = totalBoxes > 0 ? (unsentCount / totalBoxes) * 100 : 0;
        const scheduledPercent = totalBoxes > 0 ? (scheduledCount / totalBoxes) * 100 : 0;
        const sentPercent = totalBoxes > 0 ? (sentCount / totalBoxes) * 100 : 0;

        const barHtml = `
            <div class="shipping-visual-bar" title="미발송: ${unsentCount}, 발송예정: ${scheduledCount}, 발송완료: ${sentCount}">
                ${unsentPercent > 0 ? `<div class="svb-unsent" style="width: ${unsentPercent}%;"></div>` : ''}
                ${scheduledPercent > 0 ? `<div class="svb-scheduled" style="width: ${scheduledPercent}%;"></div>` : ''}
                ${sentPercent > 0 ? `<div class="svb-sent" style="width: ${sentPercent}%;"></div>` : ''}
            </div>
        `;

        const summaryHtml = `
            <div class="shipping-visual-summary">
                <span class="svs-unsent">미발송 ${unsentCount}</span>
                <span class="svs-divider">|</span>
                <span class="svs-scheduled">예정 ${scheduledCount}</span>
                <span class="svs-divider">|</span>
                <span class="svs-sent">완료 ${sentCount}</span>
            </div>
        `;

        return `<div class="shipping-visual-container">${barHtml}${summaryHtml}</div>`;
    };

    // 테이블 렌더링
    const renderTable = (ordersToRender) => {
        document.getElementById('row-count-display').textContent = `총 ${ordersToRender.length}건`;
        orderTableBody.innerHTML = '';
        if (ordersToRender.length === 0) {
            orderTableBody.innerHTML = '<tr><td colspan="17" style="text-align:center;">표시할 주문이 없습니다.</td></tr>';
            return;
        }

        const isCompact = appState.viewMode === 'compact';
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        ordersToRender.forEach((order, index) => {
            const row = document.createElement('tr');
            
            if (isOrderDelayed(order, today)) {
                row.classList.add('urgency-critical');
            }

            const senderNameHTML = isCompact ? 
                `<td class="col-name clickable-name" data-phone="${order.senderPhone}" data-address="${order.senderAddress}">${order.senderName}</td>` : 
                `<td class="col-name">${order.senderName}</td>`;

            const receiverGroupEndClass = isCompact ? ' col-group-end' : '';

            const receiverNameHTML = isCompact ? 
                `<td class="col-name clickable-name${receiverGroupEndClass}" data-phone="${order.receiverPhone}" data-address="${order.receiverAddress}">${order.receiverName}</td>` : 
                `<td class="col-name">${order.receiverName}</td>`;
            
            const receiverPhoneHTML = `<td class="col-detail">${order.receiverPhone}</td>`;
            const receiverAddressHTML = `<td class="col-detail${isCompact ? '' : ' col-group-end'}">${order.receiverAddress}</td>`;

            const totalPrice = (parseInt(order.sweetPersimmon) || 0) * appState.prices.sweet + (parseInt(order.daebongPersimmon) || 0) * appState.prices.daebong;

            row.innerHTML = `
                <td>${index + 1}</td>
                <td class="col-group-end">${formatDateWithDay(order.orderDate)}</td>
                ${senderNameHTML}
                <td class="col-detail">${order.senderPhone}</td>
                <td class="col-detail col-group-end">${order.senderAddress}</td>
                ${receiverNameHTML}
                ${receiverPhoneHTML}
                ${receiverAddressHTML}
                <td>${order.sweetPersimmon}</td>
                <td>${order.daebongPersimmon}</td>
                <td>${formatCurrency(totalPrice)}</td>
                <td class="payment-status-cell col-group-end" data-id="${order.id}">
                    <span class="payment-status ${order.paymentStatus ? 'paid' : 'unpaid'}">
                        ${order.paymentStatus ? 'O' : 'X'}
                    </span>
                </td>
                <td>
                    ${createShippingStatusVisual(order)}
                </td>
                <td class="col-group-end">
                    <button class="shipping-btn" data-id="${order.id}">배송관리</button>
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
    };

    // 데이터 필터링 및 정렬
    const getFilteredAndSortedOrders = () => {
        let processedOrders = [...appState.orders];
        
        const searchTerm = searchInput.value.toLowerCase();
        if (searchTerm) {
            processedOrders = processedOrders.filter(order => 
                Object.values(order).some(value => 
                    String(value).toLowerCase().includes(searchTerm)
                )
            );
        }

        let filteredByDashboard = processedOrders;
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        if (appState.currentFilter) {
            filteredByDashboard = processedOrders.filter(order => {
                const stats = getOrderShippingStats(order);
                const isDelayed = isOrderDelayed(order, today);

                switch (appState.currentFilter) {
                    case 'total-orders': return true;
                    case 'completed-orders':
                        const totalBoxes = stats.totalSweet + stats.totalDaebong;
                        const sentBoxes = stats.sentSweet + stats.sentDaebong;
                        return totalBoxes > 0 && totalBoxes === sentBoxes;
                    case 'sent-orders': return (stats.sentSweet + stats.sentDaebong) > 0;
                    case 'unsent-orders': return stats.unsentSweet > 0 || stats.unsentDaebong > 0;
                    case 'scheduled-orders': return stats.scheduledSweet > 0 || stats.scheduledDaebong > 0;
                    case 'delayed-orders': return isDelayed;
                    case 'total-sweet': return stats.totalSweet > 0;
                    case 'total-daebong': return stats.totalDaebong > 0;
                    case 'sent-sweet': return stats.sentSweet > 0;
                    case 'sent-daebong': return stats.sentDaebong > 0;
                    case 'unsent-sweet': return stats.unsentSweet > 0;
                    case 'unsent-daebong': return stats.unsentDaebong > 0;
                    case 'scheduled-sweet': return stats.scheduledSweet > 0;
                    case 'scheduled-daebong': return stats.scheduledDaebong > 0;
                    case 'delayed-sweet': return isDelayed && stats.unsentSweet > 0;
                    case 'delayed-daebong': return isDelayed && stats.unsentDaebong > 0;
                    default: return true;
                }
            });
        }

        filteredByDashboard.forEach(order => {
            order.totalPrice = (parseInt(order.sweetPersimmon) || 0) * appState.prices.sweet + (parseInt(order.daebongPersimmon) || 0) * appState.prices.daebong;
        });

        filteredByDashboard.sort((a, b) => {
            const column = appState.currentSort.column;
            let valA = a[column];
            let valB = b[column];

            switch (column) {
                case 'sweetPersimmon':
                case 'daebongPersimmon':
                case 'totalPrice':
                    valA = parseInt(valA) || 0;
                    valB = parseInt(valB) || 0;
                    break;
                case 'orderDate':
                    valA = new Date(valA);
                    valB = new Date(valB);
                    break;
            }

            let comparison = 0;
            if (valA > valB) comparison = 1; else if (valA < valB) comparison = -1;
            return appState.currentSort.direction === 'asc' ? comparison : -comparison;
        });

        return filteredByDashboard;
    };

    // 모달 관리
    const openOrderModal = (orderId = null) => {
        orderForm.reset();
        if (orderId) {
            const order = appState.orders.find(o => o.id == orderId);
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
        sweetPriceInput.value = appState.prices.sweet;
        daebongPriceInput.value = appState.prices.daebong;
        priceModal.style.display = 'flex';
    };
    const closePriceModal = () => { priceModal.style.display = 'none'; };

    const openAggregationModal = () => {
        renderAggregationResults();
        aggregationModal.style.display = 'flex';
    };
    const closeAggregationModal = () => { aggregationModal.style.display = 'none'; };

    // 배송 관리 모달
    const openShippingManagementModal = (orderId) => {
        const order = appState.orders.find(o => o.id == orderId);
        if (!order) return;

        shippingOrderIdInput.value = orderId;
        tempShippingDetails = JSON.parse(JSON.stringify(order.shippingDetails || { sweetPersimmon: [], daebongPersimmon: [] }));
        editingShippingItem = null;

        renderShippingModal(order);
        shippingModal.style.display = 'flex';
    };

    const closeShippingManagementModal = () => {
        shippingModal.style.display = 'none';
        editingShippingItem = null;
    };

    const renderShippingModal = (order) => {
        const sweetTotal = parseInt(order.sweetPersimmon) || 0;
        const daebongTotal = parseInt(order.daebongPersimmon) || 0;

        shippingSweetTotal.textContent = sweetTotal;
        shippingDaebongTotal.textContent = daebongTotal;

        const sweetScheduledCount = (tempShippingDetails.sweetPersimmon || [])
            .filter(item => item.status === '발송예정')
            .reduce((sum, item) => sum + item.count, 0);
        const sweetSentCount = (tempShippingDetails.sweetPersimmon || [])
            .filter(item => item.status === '발송완료')
            .reduce((sum, item) => sum + item.count, 0);

        const daebongScheduledCount = (tempShippingDetails.daebongPersimmon || [])
            .filter(item => item.status === '발송예정')
            .reduce((sum, item) => sum + item.count, 0);
        const daebongSentCount = (tempShippingDetails.daebongPersimmon || [])
            .filter(item => item.status === '발송완료')
            .reduce((sum, item) => sum + item.count, 0);

        document.getElementById('sweet-scheduled-header').textContent = `발송 예정 (${sweetScheduledCount}상자)`;
        document.getElementById('sweet-sent-header').textContent = `발송 완료 (${sweetSentCount}상자)`;
        document.getElementById('daebong-scheduled-header').textContent = `발송 예정 (${daebongScheduledCount}상자)`;
        document.getElementById('daebong-sent-header').textContent = `발송 완료 (${daebongSentCount}상자)`;

        renderShippingList('sweetPersimmon', sweetTotal);
        renderShippingList('daebongPersimmon', daebongTotal);
        
        document.querySelectorAll('.shipping-item.editing').forEach(el => el.classList.remove('editing'));
        shippingAddForms.forEach(form => {
            form.reset();
            form.querySelector('button').textContent = '추가';
        });
        editingShippingItem = null;
    };

    const renderShippingList = (type, total) => {
        const details = tempShippingDetails[type] || [];
        let scheduledHtml = '';
        let sentHtml = '';
        let currentCount = 0;

        const createLi = (item, originalIndex) => {
            const isEditing = editingShippingItem && editingShippingItem.type === type && editingShippingItem.originalIndex === originalIndex;
            const itemClass = `shipping-item status-${item.status === '발송예정' ? 'scheduled' : 'sent'} ${isEditing ? 'editing' : ''}`;
            return `<li class="${itemClass}" data-type="${type}" data-index="${originalIndex}">
                        <div class="shipping-item-info">
                            <span class="shipping-item-date">${item.date}</span>
                            <span class="shipping-item-count"><strong>${item.count}</strong>상자</span>
                        </div>
                        <button class="shipping-delete-btn" data-type="${type}" data-index="${originalIndex}">삭제</button>
                    </li>`;
        };

        details.forEach((item, index) => {
            if (item.status === '발송예정') {
                scheduledHtml += createLi(item, index);
            } else if (item.status === '발송완료') {
                sentHtml += createLi(item, index);
            }
            currentCount += item.count;
        });

        if (type === 'sweetPersimmon') {
            shippingSweetScheduledList.innerHTML = scheduledHtml;
            shippingSweetSentList.innerHTML = sentHtml;
            shippingSweetUnsent.textContent = total - currentCount;
        } else {
            shippingDaebongScheduledList.innerHTML = scheduledHtml;
            shippingDaebongSentList.innerHTML = sentHtml;
            shippingDaebongUnsent.textContent = total - currentCount;
        }
    };

    const calculateAggregations = () => {
        const results = {
            paid: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            unpaid: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            shipped: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            scheduled: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
            unshipped: { sweet: 0, daebong: 0, total: 0, sweetCount: 0, daebongCount: 0 },
        };

        appState.orders.forEach(order => {
            const stats = getOrderShippingStats(order);
            const sweetValue = stats.totalSweet * appState.prices.sweet;
            const daebongValue = stats.totalDaebong * appState.prices.daebong;

            if (order.paymentStatus) {
                results.paid.sweet += sweetValue;
                results.paid.daebong += daebongValue;
                results.paid.sweetCount += stats.totalSweet;
                results.paid.daebongCount += stats.totalDaebong;
            } else {
                results.unpaid.sweet += sweetValue;
                results.unpaid.daebong += daebongValue;
                results.unpaid.sweetCount += stats.totalSweet;
                results.unpaid.daebongCount += stats.totalDaebong;
            }

            // Shipped
            results.shipped.sweetCount += stats.sentSweet;
            results.shipped.daebongCount += stats.sentDaebong;
            results.shipped.sweet += stats.sentSweet * appState.prices.sweet;
            results.shipped.daebong += stats.sentDaebong * appState.prices.daebong;

            // Scheduled
            results.scheduled.sweetCount += stats.scheduledSweet;
            results.scheduled.daebongCount += stats.scheduledDaebong;
            results.scheduled.sweet += stats.scheduledSweet * appState.prices.sweet;
            results.scheduled.daebong += stats.scheduledDaebong * appState.prices.daebong;

            // Unshipped
            results.unshipped.sweetCount += stats.unsentSweet;
            results.unshipped.daebongCount += stats.unsentDaebong;
            results.unshipped.sweet += stats.unsentSweet * appState.prices.sweet;
            results.unshipped.daebong += stats.unsentDaebong * appState.prices.daebong;
        });

        results.paid.total = results.paid.sweet + results.paid.daebong;
        results.unpaid.total = results.unpaid.sweet + results.unpaid.daebong;
        results.shipped.total = results.shipped.sweet + results.shipped.daebong;
        results.scheduled.total = results.scheduled.sweet + results.scheduled.daebong;
        results.unshipped.total = results.unshipped.sweet + results.unshipped.daebong;

        return results;
    };

    const renderAggregationResults = () => {
        const results = calculateAggregations();
        const criteria = document.querySelector('input[name="agg-criteria"]:checked').value;

        let html = '';

        if (criteria === 'payment') {
            const data = results;
            const cat1 = 'paid';
            const cat2 = 'unpaid';
            const cat1Label = '입금 완료';
            const cat2Label = '미입금';
            
            const totalSweetCount = data[cat1].sweetCount + data[cat2].sweetCount;
            const totalDaebongCount = data[cat1].daebongCount + data[cat2].daebongCount;
            const totalSweetValue = data[cat1].sweet + data[cat2].sweet;
            const totalDaebongValue = data[cat1].daebong + data[cat2].daebong;
            const grandTotal = data[cat1].total + data[cat2].total;

            html = `
                <h3>입금여부 기준 집계</h3>
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
        } else { // shipping criteria
            const data = results;
            const cat1 = 'shipped';
            const cat2 = 'scheduled';
            const cat3 = 'unshipped';
            const cat1Label = '발송 완료';
            const cat2Label = '발송 예정';
            const cat3Label = '미발송';
            
            const totalSweetCount = data[cat1].sweetCount + data[cat2].sweetCount + data[cat3].sweetCount;
            const totalDaebongCount = data[cat1].daebongCount + data[cat2].daebongCount + data[cat3].daebongCount;
            const totalSweetValue = data[cat1].sweet + data[cat2].sweet + data[cat3].sweet;
            const totalDaebongValue = data[cat1].daebong + data[cat2].daebong + data[cat3].daebong;
            const grandTotal = data[cat1].total + data[cat2].total + data[cat3].total;

            html = `
                <h3>발송여부 기준 집계</h3>
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
                        <tr>
                            <td>${cat3Label}</td>
                            <td>${formatCurrency(data[cat3].sweet)}</td>
                            <td>${data[cat3].sweetCount}</td>
                            <td>${formatCurrency(data[cat3].daebong)}</td>
                            <td>${data[cat3].daebongCount}</td>
                            <td>${formatCurrency(data[cat3].total)}</td>
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
        }
        aggregationResults.innerHTML = html;
    };

    const saveStateToFile = () => {
        const stateToSave = {
            orders: appState.orders,
            prices: appState.prices,
            redThreshold: appState.redThreshold,
            viewMode: appState.viewMode
        };
        const blob = new Blob([JSON.stringify(stateToSave, null, 2)], { type: 'application/json' });
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
                const loadedState = JSON.parse(event.target.result);
                appState.orders = loadedState.orders || [];
                appState.prices = loadedState.prices || { sweet: 25000, daebong: 20000 };
                appState.redThreshold = loadedState.redThreshold || 7;
                appState.viewMode = loadedState.viewMode || 'compact';
                
                appState.orders.forEach(order => {
                    order.shippingDetails = order.shippingDetails || { sweetPersimmon: [], daebongPersimmon: [] };
                });

                alert('데이터를 성공적으로 불러왔습니다.');
            } catch (error) {
                alert('파일을 읽는 중 오류가 발생했습니다. 유효한 데이터 파일인지 확인해주세요.');
                console.error("Failed to parse JSON file:", error);
            }
            refreshUI();
        };
        reader.readAsText(file);
        e.target.value = '';
    };

    newOrderBtn.addEventListener('click', () => openOrderModal());
    closeOrderModalBtn.addEventListener('click', closeOrderModal);
    priceSettingsBtn.addEventListener('click', openPriceModal);
    closePriceModalBtn.addEventListener('click', closePriceModal);
    aggregationBtn.addEventListener('click', openAggregationModal);
    closeAggregationModalBtn.addEventListener('click', closeAggregationModal);
    aggregationControls.addEventListener('change', renderAggregationResults);
    saveDataBtn.addEventListener('click', saveStateToFile);
    loadDataInput.addEventListener('change', loadStateFromFile);
    closeShippingModalBtn.addEventListener('click', closeShippingManagementModal);

    window.addEventListener('click', (e) => { 
        if (e.target === orderModal) closeOrderModal();
        if (e.target === priceModal) closePriceModal();
        if (e.target === aggregationModal) closeAggregationModal();
        if (e.target === shippingModal) closeShippingManagementModal();
    });

    toggleViewBtn.addEventListener('click', () => {
        appState.viewMode = appState.viewMode === 'compact' ? 'expanded' : 'compact';
        refreshUI();
        saveStateToLocalStorage();
    });

    resetViewBtn.addEventListener('click', () => {
        searchInput.value = '';
        appState.currentFilter = null;
        document.querySelectorAll('.summary-item').forEach(item => item.classList.remove('active'));
        refreshUI();
        saveStateToLocalStorage();
    });

    dashboard.addEventListener('click', (e) => {
        const filterTarget = e.target.closest('[data-filter]');
        if (filterTarget) {
            const filterId = filterTarget.dataset.filter;
            
            if (appState.currentFilter === filterId) {
                appState.currentFilter = null; // Toggle off if clicking the same filter
            } else {
                appState.currentFilter = filterId;
            }
            
            document.querySelectorAll('[data-filter]').forEach(item => item.classList.remove('active'));
            
            if (appState.currentFilter) {
                // Highlight all elements that match the current filter
                document.querySelectorAll(`[data-filter="${appState.currentFilter}"]`).forEach(item => {
                    item.classList.add('active');
                });
            }
            refreshUI();
            saveStateToLocalStorage();
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
            const index = appState.orders.findIndex(o => o.id == id);
            if(index !== -1) appState.orders[index] = { ...appState.orders[index], ...orderData };
            refreshUI();
            closeOrderModal();
            saveStateToLocalStorage();
        } else {
            const matchingOrders = appState.orders.filter(o => 
                o.senderName === orderData.senderName && o.receiverName === orderData.receiverName
            );

            const createNewOrder = () => {
                appState.orders.push({ 
                    id: Date.now(), 
                    shippingDetails: { sweetPersimmon: [], daebongPersimmon: [] },
                    ...orderData 
                });
                refreshUI();
                closeOrderModal();
                saveStateToLocalStorage();
            };

            if (matchingOrders.length > 0) {
                let message = `동일한 이름(보내는 분: ${orderData.senderName}, 받는 분: ${orderData.receiverName})으로 등록된 주문이 ${matchingOrders.length}건 있습니다.\n\n`;
                matchingOrders.forEach(o => {
                    message += `- 주문일: ${o.orderDate}, 단감: ${o.sweetPersimmon}, 대봉: ${o.daebongPersimmon}\n`;
                });
                message += `\n그래도 추가하시겠습니까?`;

                if (confirm(message)) {
                    createNewOrder();
                }
            } else {
                createNewOrder();
            }
        }
    });

    priceSettingsForm.addEventListener('submit', (e) => {
        e.preventDefault();
        appState.prices.sweet = parseInt(sweetPriceInput.value) || 0;
        appState.prices.daebong = parseInt(daebongPriceInput.value) || 0;
        refreshUI();
        closePriceModal();
        saveStateToLocalStorage();
    });

    orderTableBody.addEventListener('click', (e) => {
        const target = e.target;
        const id = target.dataset.id;

        if (target.classList.contains('shipping-btn')) {
            openShippingManagementModal(parseInt(id));
        } else if (target.classList.contains('edit-btn')) {
            openOrderModal(parseInt(id));
        } else if (target.classList.contains('delete-btn')) {
            if (confirm('정말로 이 주문을 삭제하시겠습니까?')) {
                appState.orders = appState.orders.filter(o => o.id != id);
                refreshUI();
                saveStateToLocalStorage();
            }
        } else {
            const targetCell = target.closest('td');
            if (!targetCell) return;
            const cellId = targetCell.dataset.id;

            if (targetCell.classList.contains('payment-status-cell')) {
                const index = appState.orders.findIndex(o => o.id == cellId);
                if (index !== -1) {
                    appState.orders[index].paymentStatus = !appState.orders[index].paymentStatus;
                    refreshUI();
                    saveStateToLocalStorage();
                }
            } else if (targetCell.classList.contains('clickable-name')) {
                const phone = targetCell.dataset.phone;
                const address = targetCell.dataset.address;
                alert(`연락처: ${phone}\n주소: ${address}`);
            }
        }
    });

    shippingAddForms.forEach(form => {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const type = e.target.dataset.type;
            const countInput = e.target.querySelector('input[type="number"]');
            const dateInput = e.target.querySelector('input[type="date"]');
            const statusInput = e.target.querySelector('select');

            const count = parseInt(countInput.value);
            const date = dateInput.value;
            const status = statusInput.value;

            if (!count || !date || !status) {
                alert('수량, 날짜, 상태를 모두 입력해주세요.');
                return;
            }

            const orderId = shippingOrderIdInput.value;
            const order = appState.orders.find(o => o.id == orderId);
            const total = parseInt(order[type]) || 0;
            
            let currentTotalInDetails = 0;
            if (editingShippingItem) {
                currentTotalInDetails = (tempShippingDetails[type] || [])
                    .filter((_, i) => i !== editingShippingItem.originalIndex)
                    .reduce((sum, item) => sum + item.count, 0);
            } else {
                currentTotalInDetails = (tempShippingDetails[type] || []).reduce((sum, item) => sum + item.count, 0);
            }

            if (currentTotalInDetails + count > total) {
                alert('총 주문 수량을 초과할 수 없습니다.');
                return;
            }

            if (editingShippingItem) {
                const itemToEdit = tempShippingDetails[editingShippingItem.type][editingShippingItem.originalIndex];
                itemToEdit.count = count;
                itemToEdit.date = date;
                itemToEdit.status = status;
            } else {
                const existingEntry = (tempShippingDetails[type] || []).find(item => item.date === date && item.status === status);
                if (existingEntry) {
                    existingEntry.count += count;
                } else {
                    tempShippingDetails[type].push({ status, date, count });
                }
            }
            
            e.target.reset();
            renderShippingModal(order);
        });
    });

    const setupShippingListEventListeners = (listElement) => {
        listElement.addEventListener('click', (e) => {
            const orderId = shippingOrderIdInput.value;
            const order = appState.orders.find(o => o.id == orderId);
            const target = e.target;

            if (target.classList.contains('shipping-delete-btn')) {
                e.stopPropagation();
                const type = target.dataset.type;
                const originalIndex = parseInt(target.dataset.index);
                
                tempShippingDetails[type].splice(originalIndex, 1);
                renderShippingModal(order);

            } else if (target.closest('.shipping-item')) {
                const li = target.closest('.shipping-item');
                const type = li.dataset.type;
                const originalIndex = parseInt(li.dataset.index);
                const status = li.dataset.status;

                const item = tempShippingDetails[type][originalIndex];
                if (!item) return;

                editingShippingItem = { type, originalIndex, status };

                const form = document.querySelector(`.shipping-add-form[data-type="${type}"]`);
                form.querySelector('input[type="number"]').value = item.count;
                form.querySelector('input[type="date"]').value = item.date;
                form.querySelector('select').value = item.status;
                form.querySelector('button').textContent = '수정';

                document.querySelectorAll('.shipping-item.editing').forEach(el => el.classList.remove('editing'));
                li.classList.add('editing');
            }
        });
    };

    setupShippingListEventListeners(shippingSweetScheduledList);
    setupShippingListEventListeners(shippingSweetSentList);
    setupShippingListEventListeners(shippingDaebongScheduledList);
    setupShippingListEventListeners(shippingDaebongSentList);

    shippingModalSaveBtn.addEventListener('click', () => {
        const orderId = shippingOrderIdInput.value;
        const index = appState.orders.findIndex(o => o.id == orderId);
        if (index !== -1) {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            ['sweetPersimmon', 'daebongPersimmon'].forEach(type => {
                const details = tempShippingDetails[type] || [];
                const updatedDetails = [];
                const itemsToConvert = [];

                details.forEach(item => {
                    const itemDate = new Date(item.date + 'T00:00:00');
                    if (item.status === '발송예정' && itemDate < today) {
                        itemsToConvert.push({ ...item, status: '발송완료' });
                    } else {
                        updatedDetails.push(item);
                    }
                });

                itemsToConvert.forEach(convertedItem => {
                    const existing = updatedDetails.find(i => i.date === convertedItem.date && i.status === '발송완료');
                    if (existing) {
                        existing.count += convertedItem.count;
                    } else {
                        updatedDetails.push(convertedItem);
                    }
                });
                tempShippingDetails[type] = updatedDetails;
            });

            tempShippingDetails.sweetPersimmon.sort((a, b) => new Date(a.date) - new Date(b.date));
            tempShippingDetails.daebongPersimmon.sort((a, b) => new Date(a.date) - new Date(b.date));
            appState.orders[index].shippingDetails = JSON.parse(JSON.stringify(tempShippingDetails));
        }
        
        closeShippingManagementModal();
        refreshUI();
        saveStateToLocalStorage();
    });


    searchInput.addEventListener('input', () => renderTable(getFilteredAndSortedOrders()));

    document.querySelectorAll('th[data-sort]').forEach(header => {
        header.addEventListener('click', () => {
            const column = header.dataset.sort;
            if (appState.currentSort.column === column) {
                appState.currentSort.direction = appState.currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                appState.currentSort.column = column;
                appState.currentSort.direction = 'desc';
            }
            refreshUI();
        });
    });

    redThresholdInput.addEventListener('change', () => {
        appState.redThreshold = parseInt(redThresholdInput.value) || 7;
        refreshUI();
        saveStateToLocalStorage();
    });

        exportScheduledBtn.addEventListener('click', () => {

            const itemsToExport = [];

            let itemIndex = 1;

    

            appState.orders.forEach(order => {

                const details = order.shippingDetails || { sweetPersimmon: [], daebongPersimmon: [] };

    

                (details.sweetPersimmon || []).forEach(item => {

                    if (item.status === '발송예정') {

                        itemsToExport.push({

                            '순번': itemIndex++,

                            '발송예정일': item.date,

                            '보내는분 이름': order.senderName,

                            '보내는분 연락처': order.senderPhone,

                            '보내는분 주소': order.senderAddress,

                            '받는분': order.receiverName,

                            '받는 분 연락처': order.receiverPhone,

                            '받는 분 주소': order.receiverAddress,

                            '품목': '단감',

                            '수량': item.count

                        });

                    }

                });

    

                (details.daebongPersimmon || []).forEach(item => {

                    if (item.status === '발송예정') {

                        itemsToExport.push({

                            '순번': itemIndex++,

                            '발송예정일': item.date,

                            '보내는분 이름': order.senderName,

                            '보내는분 연락처': order.senderPhone,

                            '보내는분 주소': order.senderAddress,

                            '받는분': order.receiverName,

                            '받는 분 연락처': order.receiverPhone,

                            '받는 분 주소': order.receiverAddress,

                            '품목': '대봉',

                            '수량': item.count

                        });

                    }

                });

            });

    

            if (itemsToExport.length === 0) {

                alert('발송예정인 주문이 없습니다.');

                return;

            }

    

            itemsToExport.sort((a, b) => new Date(a['발송예정일']) - new Date(b['발송예정일']));

    

            const ws_cols = [

                { wch: 8 }, 
                { wch: 15 }, 
                { wch: 12 }, 
                { wch: 15 }, 
                { wch: 30 }, 
                { wch: 12 }, 
                { wch: 15 }, 
                { wch: 30 }, 
                { wch: 8 }, 
                { wch: 8 }
            ];

    

            const worksheet = XLSX.utils.json_to_sheet(itemsToExport);

            worksheet['!cols'] = ws_cols;

    

            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(workbook, worksheet, '발송예정 목록');

    

            const today = new Date().toISOString().split('T')[0];

            XLSX.writeFile(workbook, `발송예정_목록_${today}.xlsx`);

        });
    
    const saveStateToLocalStorage = () => {
        localStorage.setItem('gamAppState', JSON.stringify(appState));
    };

    const loadStateFromLocalStorage = () => {
        const savedState = localStorage.getItem('gamAppState');
        if (savedState) {
            try {
                const loadedState = JSON.parse(savedState);
                appState.orders = loadedState.orders || [];
                appState.prices = loadedState.prices || { sweet: 25000, daebong: 20000 };
                appState.redThreshold = loadedState.redThreshold || 7;
                appState.viewMode = loadedState.viewMode || 'compact';
                appState.currentSort = loadedState.currentSort || { column: 'orderDate', direction: 'desc' };
                appState.currentFilter = loadedState.currentFilter || null;


                appState.orders.forEach(order => {
                    order.shippingDetails = order.shippingDetails || { sweetPersimmon: [], daebongPersimmon: [] };
                });

            } catch (error) {
                console.error("Failed to parse app state from local storage:", error);
                // Reset to default state if parsing fails
                appState.orders = [];
                appState.redThreshold = 7;
                appState.viewMode = 'compact';
                appState.prices = { sweet: 25000, daebong: 20000 };
                appState.currentSort = { column: 'orderDate', direction: 'desc' };
                appState.currentFilter = null;
            }
        }
    };

    const updateScheduledToSent = () => {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        appState.orders.forEach(order => {
            if (!order.shippingDetails) return;

            ['sweetPersimmon', 'daebongPersimmon'].forEach(type => {
                const updatedDetails = [];
                const itemsToConvert = [];

                order.shippingDetails[type].forEach(item => {
                    const itemDate = new Date(item.date + 'T00:00:00');
                    if (item.status === '발송예정' && itemDate < today) {
                        itemsToConvert.push({ ...item, status: '발송완료' });
                    } else {
                        updatedDetails.push(item);
                    }
                });

                itemsToConvert.forEach(convertedItem => {
                    const existing = updatedDetails.find(i => i.date === convertedItem.date && i.status === '발송완료');
                    if (existing) {
                        existing.count += convertedItem.count;
                    } else {
                        updatedDetails.push(convertedItem);
                    }
                });
                order.shippingDetails[type] = updatedDetails;
            });
        });
    };

    const initialize = () => {
        loadStateFromLocalStorage();
        updateScheduledToSent();
        refreshUI();
        if (appState.currentFilter) {
            document.querySelectorAll('.summary-item').forEach(item => item.classList.remove('active'));
            const activeSpan = document.getElementById(appState.currentFilter);
            if (activeSpan) {
                activeSpan.closest('.summary-item').classList.add('active');
            }
        }
    };

    initialize();
    window.addEventListener('beforeunload', saveStateToLocalStorage);
});
