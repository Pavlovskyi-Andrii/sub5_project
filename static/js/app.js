// Road to SUB5 Dashboard JavaScript

// Global state
let charts = {};

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    initializeTabs();
    initializeSyncButton();
    initializeFilters();
    loadDashboardData();

    // Auto-refresh every 5 minutes
    setInterval(loadDashboardData, 5 * 60 * 1000);
});

// Tab Management
function initializeTabs() {
    const tabs = document.querySelectorAll('.tab');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const targetTab = this.getAttribute('data-tab');

            // Remove active class from all tabs and panes
            tabs.forEach(t => t.classList.remove('active'));
            tabPanes.forEach(pane => pane.classList.remove('active'));

            // Add active class to clicked tab and corresponding pane
            this.classList.add('active');
            document.getElementById(targetTab).classList.add('active');

            // Load data for specific tab
            if (targetTab === 'activities') {
                loadActivities();
            } else if (targetTab === 'weekly') {
                loadWeeklyStats();
            } else if (targetTab === 'logs') {
                loadSyncLogs();
            }
        });
    });
}

// Sync Button
function initializeSyncButton() {
    const syncBtn = document.getElementById('syncBtn');

    syncBtn.addEventListener('click', async function() {
        if (this.classList.contains('syncing')) {
            return; // Prevent multiple clicks
        }

        this.classList.add('syncing');
        this.innerHTML = '<i class="fas fa-sync fa-spin"></i> Синхронизация...';
        this.disabled = true;

        try {
            const response = await axios.post('/api/sync');

            if (response.data.status === 'started') {
                showNotification('Синхронизация запущена в фоновом режиме', 'success');

                // Wait a bit and reload data
                setTimeout(() => {
                    loadDashboardData();
                }, 5000);
            }
        } catch (error) {
            console.error('Sync error:', error);
            showNotification('Ошибка синхронизации: ' + (error.response?.data?.message || error.message), 'error');
        } finally {
            this.classList.remove('syncing');
            this.innerHTML = '<i class="fas fa-sync"></i> Синхронизировать';
            this.disabled = false;
        }
    });
}

// Initialize Filters
function initializeFilters() {
    const applyFiltersBtn = document.getElementById('applyFilters');

    if (applyFiltersBtn) {
        applyFiltersBtn.addEventListener('click', function() {
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const activityType = document.getElementById('activityType').value;

            let url = '/api/activities?limit=100';

            if (startDate) {
                url += `&start_date=${startDate}`;
            }
            if (endDate) {
                url += `&end_date=${endDate}`;
            }
            if (activityType) {
                url += `&type=${activityType}`;
            }

            axios.get(url)
                .then(response => {
                    displayActivities(response.data);
                })
                .catch(error => {
                    console.error('Error loading filtered activities:', error);
                    showNotification('Ошибка загрузки активностей', 'error');
                });
        });
    }
}

// Load all dashboard data
async function loadDashboardData() {
    try {
        showLoading(true);

        // Load summary
        const summaryResponse = await axios.get('/api/summary');
        updateSummaryCards(summaryResponse.data);

        // Load activities for chart
        const activitiesResponse = await axios.get('/api/activities?limit=50');
        updateOverviewCharts(activitiesResponse.data);

        showLoading(false);
    } catch (error) {
        console.error('Error loading dashboard:', error);
        showLoading(false);
        showNotification('Ошибка загрузки данных', 'error');
    }
}

// Update summary cards
function updateSummaryCards(data) {
    const { week_stats, last_sync } = data;

    // Update cycling
    document.getElementById('cyclingKm').textContent =
        (week_stats.total_cycling_km || 0).toFixed(1) + ' км';

    // Update running
    document.getElementById('runningKm').textContent =
        (week_stats.total_running_km || 0).toFixed(1) + ' км';

    // Update total activities
    document.getElementById('totalActivities').textContent =
        week_stats.total_activities || 0;

    // Update last sync
    if (last_sync) {
        const syncDate = new Date(last_sync.date);
        document.getElementById('lastSync').textContent = formatDateTime(syncDate);
        document.getElementById('syncStatus').textContent =
            last_sync.status === 'success' ?
            `${last_sync.activities_synced} активностей` :
            'Ошибка';
    }
}

// Update overview charts
function updateOverviewCharts(activities) {
    // Group activities by week
    const weeklyData = groupActivitiesByWeek(activities);

    // Create/update weekly volume chart
    createWeeklyVolumeChart(weeklyData);

    // Create/update activity type chart
    createActivityTypeChart(activities);
}

// Group activities by week
function groupActivitiesByWeek(activities) {
    const weeks = {};

    activities.forEach(activity => {
        const date = new Date(activity.date);
        const weekStart = getWeekStart(date);
        const weekKey = weekStart.toISOString().split('T')[0];

        if (!weeks[weekKey]) {
            weeks[weekKey] = {
                date: weekKey,
                cycling: 0,
                running: 0,
                total: 0
            };
        }

        const distance = (activity.distance || 0) / 1000; // Convert to km

        if (activity.type && activity.type.toLowerCase().includes('cycling')) {
            weeks[weekKey].cycling += distance;
        } else if (activity.type && activity.type.toLowerCase().includes('running')) {
            weeks[weekKey].running += distance;
        }

        weeks[weekKey].total += distance;
    });

    return Object.values(weeks).sort((a, b) => new Date(a.date) - new Date(b.date));
}

// Get week start (Monday)
function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(d.setDate(diff));
}

// Create weekly volume chart
function createWeeklyVolumeChart(weeklyData) {
    const ctx = document.getElementById('weeklyVolumeChart');
    if (!ctx) return;

    // Destroy existing chart
    if (charts.weeklyVolume) {
        charts.weeklyVolume.destroy();
    }

    charts.weeklyVolume = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: weeklyData.map(w => formatDate(new Date(w.date))),
            datasets: [
                {
                    label: 'Велосипед',
                    data: weeklyData.map(w => w.cycling),
                    backgroundColor: 'rgba(54, 162, 235, 0.6)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Бег',
                    data: weeklyData.map(w => w.running),
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Километры'
                    }
                }
            }
        }
    });
}

// Create activity type chart
function createActivityTypeChart(activities) {
    const ctx = document.getElementById('activityTypeChart');
    if (!ctx) return;

    // Count by type
    const types = {};
    activities.forEach(activity => {
        const type = activity.type || 'Неизвестно';
        types[type] = (types[type] || 0) + 1;
    });

    // Destroy existing chart
    if (charts.activityType) {
        charts.activityType.destroy();
    }

    charts.activityType = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(types),
            datasets: [{
                data: Object.values(types),
                backgroundColor: [
                    'rgba(54, 162, 235, 0.8)',
                    'rgba(255, 99, 132, 0.8)',
                    'rgba(255, 206, 86, 0.8)',
                    'rgba(75, 192, 192, 0.8)',
                    'rgba(153, 102, 255, 0.8)'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    });
}

// Load activities list
async function loadActivities() {
    try {
        const response = await axios.get('/api/activities?limit=100');
        displayActivities(response.data);
    } catch (error) {
        console.error('Error loading activities:', error);
        showNotification('Ошибка загрузки активностей', 'error');
    }
}

// Display activities in table
function displayActivities(activities) {
    const tbody = document.getElementById('activitiesTableBody');

    if (activities.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" class="no-data">Нет данных о тренировках</td></tr>';
        return;
    }

    let html = '';

    activities.forEach(activity => {
        const isCycling = activity.type && activity.type.toLowerCase().includes('cycling');
        const isRunning = activity.type && activity.type.toLowerCase().includes('running');

        html += `
            <tr>
                <td>${formatDate(new Date(activity.date))}</td>
                <td>${activity.name || '—'}</td>
                <td><span class="activity-type">${activity.type || '—'}</span></td>
                <td>${activity.duration ? formatDuration(activity.duration) : '—'}</td>
                <td>${activity.distance ? (activity.distance / 1000).toFixed(2) + ' км' : '—'}</td>
                <td>${activity.avg_speed ? (activity.avg_speed * 3.6).toFixed(1) + ' км/ч' : '—'}</td>
                <td>${activity.avg_hr || '—'}</td>
                <td>${isCycling && activity.avg_power ? activity.avg_power : '—'}</td>
                <td>${isCycling && activity.normalized_power ? activity.normalized_power : '—'}</td>
                <td>${isCycling && activity.avg_cadence ? activity.avg_cadence : '—'}</td>
                <td>${isCycling && activity.tss ? activity.tss : '—'}</td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
}

// Load weekly statistics
async function loadWeeklyStats() {
    try {
        const response = await axios.get('/api/weekly-stats');
        displayWeeklyStats(response.data);
    } catch (error) {
        console.error('Error loading weekly stats:', error);
        showNotification('Ошибка загрузки статистики', 'error');
    }
}

// Display weekly statistics
function displayWeeklyStats(stats) {
    const tbody = document.getElementById('weeklyStatsTableBody');

    if (stats.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="no-data">Нет данных о недельной статистике</td></tr>';
        return;
    }

    let html = '';

    stats.forEach(week => {
        html += `
            <tr>
                <td>${formatDate(new Date(week.week_start))} - ${formatDate(new Date(week.week_end))}</td>
                <td><strong>${week.total_cycling_km ? week.total_cycling_km.toFixed(1) : '0.0'}</strong></td>
                <td>${week.total_cycling_time ? formatDuration(week.total_cycling_time) : '—'}</td>
                <td><strong>${week.total_running_km ? week.total_running_km.toFixed(1) : '0.0'}</strong></td>
                <td>${week.total_running_time ? formatDuration(week.total_running_time) : '—'}</td>
                <td>${week.total_activities || 0}</td>
                <td>${week.avg_hrv ? week.avg_hrv.toFixed(0) : '—'}</td>
            </tr>
        `;
    });

    tbody.innerHTML = html;

    // Create charts for weekly stats
    createWeeklyDistanceChart(stats);
    createWeeklyTimeChart(stats);
}

// Create weekly distance chart
function createWeeklyDistanceChart(stats) {
    const ctx = document.getElementById('weeklyDistanceChart');
    if (!ctx) return;

    // Destroy existing chart
    if (charts.weeklyDistance) {
        charts.weeklyDistance.destroy();
    }

    const labels = stats.map(w => formatDate(new Date(w.week_start)));

    charts.weeklyDistance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Велосипед (км)',
                    data: stats.map(w => w.total_cycling_km || 0),
                    backgroundColor: 'rgba(54, 162, 235, 0.6)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Бег (км)',
                    data: stats.map(w => w.total_running_km || 0),
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Километры'
                    }
                }
            }
        }
    });
}

// Create weekly time chart
function createWeeklyTimeChart(stats) {
    const ctx = document.getElementById('weeklyTimeChart');
    if (!ctx) return;

    // Destroy existing chart
    if (charts.weeklyTime) {
        charts.weeklyTime.destroy();
    }

    const labels = stats.map(w => formatDate(new Date(w.week_start)));

    charts.weeklyTime = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Велосипед (часы)',
                    data: stats.map(w => (w.total_cycling_time || 0) / 3600),
                    backgroundColor: 'rgba(54, 162, 235, 0.6)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Бег (часы)',
                    data: stats.map(w => (w.total_running_time || 0) / 3600),
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Часы'
                    },
                    ticks: {
                        callback: function(value) {
                            return value.toFixed(1) + ' ч';
                        }
                    }
                }
            }
        }
    });
}

// Load sync logs
async function loadSyncLogs() {
    try {
        const response = await axios.get('/api/sync-logs');
        displaySyncLogs(response.data);
    } catch (error) {
        console.error('Error loading sync logs:', error);
        showNotification('Ошибка загрузки логов', 'error');
    }
}

// Display sync logs
function displaySyncLogs(logs) {
    const container = document.getElementById('syncLogsList');

    if (logs.length === 0) {
        container.innerHTML = '<p class="no-data">Нет логов синхронизации</p>';
        return;
    }

    let html = '<div class="logs-container">';

    logs.forEach(log => {
        const statusClass = log.status === 'success' ? 'success' : 'error';
        html += `
            <div class="log-entry ${statusClass}">
                <div class="log-header">
                    <span class="log-time">${formatDateTime(new Date(log.sync_date))}</span>
                    <span class="log-status ${statusClass}">${log.status}</span>
                </div>
                <div class="log-details">
                    <p>Синхронизировано активностей: ${log.activities_synced}</p>
                    ${log.error_message ? `<p class="error-message">Ошибка: ${log.error_message}</p>` : ''}
                </div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
}

// Helper: Format date
function formatDate(date) {
    return date.toLocaleDateString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

// Helper: Format datetime
function formatDateTime(date) {
    return date.toLocaleString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

// Helper: Format duration (seconds to HH:MM:SS)
function formatDuration(seconds) {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = Math.floor(seconds % 60);

    if (hours > 0) {
        return `${hours}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    } else {
        return `${minutes}:${secs.toString().padStart(2, '0')}`;
    }
}

// Show/hide loading overlay
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = show ? 'flex' : 'none';
    }
}

// Show notification
function showNotification(message, type = 'info') {
    const notification = document.getElementById('notification');
    const notificationText = document.getElementById('notificationText');

    if (notification && notificationText) {
        notificationText.textContent = message;
        notification.className = `notification ${type} show`;

        setTimeout(() => {
            notification.classList.remove('show');
        }, 5000);
    }
}
