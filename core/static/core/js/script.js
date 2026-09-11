/* ============================================================
   LALANTHA DISTRIBUTORS DMS - Main JavaScript
   ============================================================ */

// ===== CSRF TOKEN HELPER =====
function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== '') {
        const cookies = document.cookie.split(';');
        for (let i = 0; i < cookies.length; i++) {
            const cookie = cookies[i].trim();
            if (cookie.substring(0, name.length + 1) === (name + '=')) {
                cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                break;
            }
        }
    }
    return cookieValue;
}

const CSRF_TOKEN = getCookie('csrftoken');

// ===== AUTO-DISMISS MESSAGES =====
document.addEventListener('DOMContentLoaded', function() {
    const messages = document.querySelectorAll('.msg');
    messages.forEach(msg => {
        setTimeout(() => {
            msg.style.transition = 'opacity 0.5s, transform 0.5s';
            msg.style.opacity = '0';
            msg.style.transform = 'translateX(100%)';
            setTimeout(() => msg.remove(), 500);
        }, 5000); // Auto-dismiss after 5 seconds
    });
    
    // ===== SIDEBAR TOGGLE (Mobile) =====
    const menuToggle = document.querySelector('.menu-toggle');
    const sidebar = document.querySelector('.sidebar');
    
    if (menuToggle && sidebar) {
        menuToggle.addEventListener('click', function() {
            sidebar.classList.toggle('open');
        });
        
        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', function(e) {
            if (window.innerWidth <= 768 && 
                sidebar.classList.contains('open') &&
                !sidebar.contains(e.target) &&
                !menuToggle.contains(e.target)) {
                sidebar.classList.remove('open');
            }
        });
    }
});

// ===== FORMAT CURRENCY =====
function formatRs(amount) {
    return 'Rs ' + parseFloat(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ===== CONFIRM DELETE HELPER =====
function confirmDelete(message) {
    return confirm(message || 'Are you sure you want to delete this?');
}

// ===== GLOBAL ERROR HANDLER =====
window.addEventListener('error', function(e) {
    console.error('Global error:', e.message, e.filename, e.lineno);
});