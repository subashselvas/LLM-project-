// Main JavaScript File

// Initialize when document is ready
$(document).ready(function() {
    // Auto-hide alerts after 5 seconds
    setTimeout(function() {
        $('.alert').fadeOut('slow');
    }, 5000);

    // Initialize tooltips
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    var tooltipList = tooltipTriggerList.map(function(tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    // Initialize popovers
    var popoverTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="popover"]'));
    var popoverList = popoverTriggerList.map(function(popoverTriggerEl) {
        return new bootstrap.Popover(popoverTriggerEl);
    });

    // Password strength indicator
    $('#password').on('keyup', function() {
        var password = $(this).val();
        var strength = calculatePasswordStrength(password);
        
        var strengthBar = $('#password-strength');
        if (strengthBar.length) {
            strengthBar.css('width', strength.percentage + '%');
            strengthBar.removeClass('bg-danger bg-warning bg-success');
            strengthBar.addClass(strength.class);
            strengthBar.text(strength.text);
        }
    });

    // Confirm delete actions
    $('.delete-confirm').on('click', function(e) {
        e.preventDefault();
        var form = $(this).closest('form');
        
        if (confirm('Are you sure you want to delete this item? This action cannot be undone.')) {
            form.submit();
        }
    });

    // Dynamic search
    var searchTimeout;
    $('#search-input').on('keyup', function() {
        clearTimeout(searchTimeout);
        var query = $(this).val();
        
        searchTimeout = setTimeout(function() {
            if (query.length > 2) {
                performSearch(query);
            }
        }, 500);
    });

    // Load dashboard stats
    if ($('#dashboard-stats').length) {
        loadDashboardStats();
    }
});

// Calculate password strength
function calculatePasswordStrength(password) {
    var strength = {
        percentage: 0,
        class: 'bg-danger',
        text: 'Weak'
    };

    if (password.length === 0) {
        return strength;
    }

    var score = 0;

    // Length check
    if (password.length >= 8) score += 25;
    else if (password.length >= 6) score += 15;

    // Uppercase check
    if (/[A-Z]/.test(password)) score += 25;

    // Lowercase check
    if (/[a-z]/.test(password)) score += 25;

    // Number check
    if (/[0-9]/.test(password)) score += 25;

    // Special character check
    if (/[^A-Za-z0-9]/.test(password)) score += 25;

    // Cap at 100
    score = Math.min(score, 100);
    strength.percentage = score;

    // Determine strength level
    if (score >= 80) {
        strength.class = 'bg-success';
        strength.text = 'Strong';
    } else if (score >= 50) {
        strength.class = 'bg-warning';
        strength.text = 'Medium';
    } else {
        strength.class = 'bg-danger';
        strength.text = 'Weak';
    }

    return strength;
}

// Perform search
function performSearch(query) {
    var searchType = $('#search-type').val();
    var resultsContainer = $('#search-results');
    
    if (!resultsContainer.length) return;

    // Show loading spinner
    showLoading();

    $.ajax({
        url: '/api/search',
        method: 'GET',
        data: {
            type: searchType,
            q: query
        },
        success: function(data) {
            displaySearchResults(data);
        },
        error: function(xhr, status, error) {
            console.error('Search failed:', error);
            showNotification('Search failed. Please try again.', 'error');
        },
        complete: function() {
            hideLoading();
        }
    });
}

// Display search results
function displaySearchResults(results) {
    var container = $('#search-results');
    container.empty();

    if (!results || results.length === 0) {
        container.html('<p class="text-muted text-center">No results found</p>');
        return;
    }

    var table = $('<table class="table table-hover"></table>');
    var thead = $('<thead><tr></tr></thead>');
    var tbody = $('<tbody></tbody>');

    // Add headers based on first result
    var firstItem = results[0];
    var headers = Object.keys(firstItem);
    
    headers.forEach(function(header) {
        thead.find('tr').append('<th>' + header.replace('_', ' ').toUpperCase() + '</th>');
    });
    thead.find('tr').append('<th>Actions</th>');

    // Add data rows
    results.forEach(function(item) {
        var row = $('<tr></tr>');
        
        headers.forEach(function(header) {
            row.append('<td>' + (item[header] || '-') + '</td>');
        });
        
        // Add action buttons
        var actions = $('<td></td>');
        if (item.id) {
            actions.append('<button class="btn btn-sm btn-info me-1" onclick="viewItem(' + item.id + ')">View</button>');
            actions.append('<button class="btn btn-sm btn-warning" onclick="editItem(' + item.id + ')">Edit</button>');
        }
        row.append(actions);
        
        tbody.append(row);
    });

    table.append(thead).append(tbody);
    container.append(table);
}

// Show loading spinner
function showLoading() {
    $('.spinner-overlay').addClass('show');
}

// Hide loading spinner
function hideLoading() {
    $('.spinner-overlay').removeClass('show');
}

// Show notification
function showNotification(message, type) {
    var notification = $('<div class="alert alert-' + type + ' alert-dismissible fade show" role="alert">' +
        message +
        '<button type="button" class="btn-close" data-bs-dismiss="alert"></button>' +
        '</div>');
    
    $('.notifications-container').append(notification);
    
    setTimeout(function() {
        notification.fadeOut('slow', function() {
            $(this).remove();
        });
    }, 5000);
}

// Load dashboard statistics
function loadDashboardStats() {
    $.ajax({
        url: '/api/stats/dashboard',
        method: 'GET',
        success: function(stats) {
            updateDashboardStats(stats);
        },
        error: function(xhr, status, error) {
            console.error('Failed to load stats:', error);
        }
    });
}

// Update dashboard statistics
function updateDashboardStats(stats) {
    for (var key in stats) {
        var element = $('#stat-' + key);
        if (element.length) {
            element.text(stats[key]);
            element.addClass('fade-in');
        }
    }
}

// Export to CSV
function exportToCSV(data, filename) {
    var csv = '';
    
    // Add headers
    var headers = Object.keys(data[0]);
    csv += headers.join(',') + '\n';
    
    // Add data
    data.forEach(function(row) {
        var values = headers.map(function(header) {
            var value = row[header] || '';
            // Escape quotes and wrap in quotes if contains comma
            if (value.toString().includes(',')) {
                value = '"' + value.replace(/"/g, '""') + '"';
            }
            return value;
        });
        csv += values.join(',') + '\n';
    });
    
    // Download
    var blob = new Blob([csv], { type: 'text/csv' });
    var url = window.URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename + '.csv';
    a.click();
    window.URL.revokeObjectURL(url);
}

// Print table
function printTable(tableId) {
    var printContents = document.getElementById(tableId).outerHTML;
    var originalContents = document.body.innerHTML;
    
    document.body.innerHTML = printContents;
    window.print();
    document.body.innerHTML = originalContents;
    location.reload();
}

// Format date
function formatDate(dateString) {
    var options = { year: 'numeric', month: 'short', day: 'numeric' };
    return new Date(dateString).toLocaleDateString('en-US', options);
}

// Format currency
function formatCurrency(amount) {
    return '$' + parseFloat(amount).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

// Calculate GPA
function calculateGPA(grades) {
    var gradePoints = {
        'A': 4.0, 'A-': 3.7,
        'B+': 3.3, 'B': 3.0, 'B-': 2.7,
        'C+': 2.3, 'C': 2.0, 'C-': 1.7,
        'D+': 1.3, 'D': 1.0, 'F': 0.0
    };
    
    var totalPoints = 0;
    var count = 0;
    
    grades.forEach(function(grade) {
        if (gradePoints[grade]) {
            totalPoints += gradePoints[grade];
            count++;
        }
    });
    
    return count > 0 ? (totalPoints / count).toFixed(2) : '0.00';
}

// Handle form validation
function validateForm(formId) {
    var isValid = true;
    var form = document.getElementById(formId);
    
    if (!form) return false;
    
    var inputs = form.querySelectorAll('[required]');
    inputs.forEach(function(input) {
        if (!input.value.trim()) {
            input.classList.add('is-invalid');
            isValid = false;
        } else {
            input.classList.remove('is-invalid');
        }
    });
    
    return isValid;
}

// Auto-resize textarea
function autoResizeTextarea(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = textarea.scrollHeight + 'px';
}

// Debounce function
function debounce(func, wait) {
    var timeout;
    return function executedFunction() {
        var context = this;
        var args = arguments;
        
        var later = function() {
            timeout = null;
            func.apply(context, args);
        };
        
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}
// Fix DataTables reinitialisation warning
$.fn.dataTable.ext.errMode = 'none';