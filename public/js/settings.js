/**
 * Settings Module
 * Global settings, email settings, and signature generation
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Global Settings Modal
  // ============================================================
  
  window.openGlobalSettings = function() {
    openModal('globalSettingsModal');
    loadEmailLinksFromSettings();
    loadUserSignature();
  };
  
  window.closeGlobalSettings = function() {
    closeModal('globalSettingsModal');
  };
  
  // ============================================================
  // Email Links Settings
  // ============================================================
  
  window.loadEmailLinksFromSettings = function() {
    google.script.run.withSuccessHandler(function(settings) {
      if (settings) {
        document.getElementById('settingsBookingLink').value = settings.bookingLink || '';
        document.getElementById('settingsReviewLink').value = settings.reviewLink || '';
        document.getElementById('settingsReferralLink').value = settings.referralLink || '';
      }
    }).getGlobalSettings();
  };
  
  window.openEmailSettings = function() {
    loadEmailLinksFromSettings();
    openModal('emailSettingsModal');
  };
  
  window.closeEmailSettings = function() {
    closeModal('emailSettingsModal');
  };
  
  window.saveEmailSettings = function() {
    const settings = {
      bookingLink: document.getElementById('settingsBookingLink').value,
      reviewLink: document.getElementById('settingsReviewLink').value,
      referralLink: document.getElementById('settingsReferralLink').value
    };
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        closeEmailSettings();
        showAlert('Success', 'Settings saved successfully', 'success');
      } else {
        showAlert('Error', result.error || 'Failed to save settings', 'error');
      }
    }).saveGlobalSettings(settings);
  };
  
  // ============================================================
  // User Signature
  // ============================================================
  
  window.loadUserSignature = function() {
    google.script.run.withSuccessHandler(function(signature) {
      const preview = document.getElementById('signaturePreview');
      if (preview && signature) {
        preview.innerHTML = signature;
      }
    }).getUserSignature();
  };
  
  window.updateSignatureUserInfo = function() {
    const firstName = document.getElementById('sigFirstName')?.value || '';
    const lastName = document.getElementById('sigLastName')?.value || '';
    const title = document.getElementById('sigTitle')?.value || '';
    const phone = document.getElementById('sigPhone')?.value || '';
    const mobile = document.getElementById('sigMobile')?.value || '';
    
    generateSignaturePreview({ firstName, lastName, title, phone, mobile });
  };
  
  window.regenerateSignature = function() {
    const firstName = document.getElementById('sigFirstName')?.value || '';
    const lastName = document.getElementById('sigLastName')?.value || '';
    const title = document.getElementById('sigTitle')?.value || '';
    const phone = document.getElementById('sigPhone')?.value || '';
    const mobile = document.getElementById('sigMobile')?.value || '';
    
    const signatureHtml = generateSignatureHtml({ firstName, lastName, title, phone, mobile });
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        document.getElementById('signaturePreview').innerHTML = signatureHtml;
        showAlert('Success', 'Signature updated', 'success');
      }
    }).saveUserSignature(signatureHtml);
  };
  
  window.generateSignaturePreview = function(userInfo) {
    const preview = document.getElementById('signaturePreview');
    if (preview) {
      preview.innerHTML = generateSignatureHtml(userInfo);
    }
  };
  
  function generateSignatureHtml(info) {
    return `
      <table style="font-family: Arial, sans-serif; font-size: 12px; color: #333;">
        <tr>
          <td style="padding-right: 15px; border-right: 2px solid #BB9934;">
            <img src="https://img1.wsimg.com/isteam/ip/2c5f94ee-4964-4e9b-9b9c-a55121f8611b/WEB_Stellaris_Primary%20Logo_Horizontal_Full%20Col.png/:/rs=w:200" alt="Stellaris" style="width: 150px;">
          </td>
          <td style="padding-left: 15px;">
            <div style="font-weight: bold; font-size: 14px; color: #2C2622;">${escapeHtml(info.firstName)} ${escapeHtml(info.lastName)}</div>
            <div style="color: #666; font-size: 11px;">${escapeHtml(info.title)}</div>
            <div style="margin-top: 8px;">
              ${info.phone ? `<div>P: ${escapeHtml(info.phone)}</div>` : ''}
              ${info.mobile ? `<div>M: ${escapeHtml(info.mobile)}</div>` : ''}
            </div>
          </td>
        </tr>
      </table>
    `;
  }
  
  window.useGeneratedSignature = function() {
    const preview = document.getElementById('signaturePreview');
    if (preview) {
      const signatureHtml = preview.innerHTML;
      
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          showAlert('Success', 'Signature saved', 'success');
        }
      }).saveUserSignature(signatureHtml);
    }
  };
  
})();
