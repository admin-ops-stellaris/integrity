/**
 * settings.js - Settings and Email Configuration Module
 * 
 * Handles team-wide settings, email template links, and signature generation.
 * Email links and signatures are saved to Airtable Settings table for team-wide sync.
 * 
 * Dependencies: shared-state.js, modal-utils.js
 * 
 * Exposed to window:
 * - EMAIL_LINKS, DEFAULT_EMAIL_LINKS, SETTINGS_KEYS (config)
 * - userSignature, emailSettingsLoaded, currentUserProfile, generatedSignatureHtml (state)
 * - loadEmailLinksFromSettings, loadUserSignature
 * - openEmailSettings, closeEmailSettings, saveEmailSettings, openGlobalSettings
 * - updateSignatureUserInfo, regenerateSignature, generateSignaturePreview
 * - copySignatureForGmail, copySignatureForMercury, useGeneratedSignature
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;

  const DEFAULT_EMAIL_LINKS = {
    officeMap: 'https://maps.app.goo.gl/qm2ohJP2j1t6GqCt9',
    ourTeam: 'https://stellaris.loans/our-team',
    factFind: 'https://drive.google.com/file/d/1_U6kKck5IA3TBtFdJEzyxs_XpvcKvg9s/view?usp=sharing',
    myGov: 'https://my.gov.au/',
    myGovVideo: 'https://www.youtube.com/watch?v=bSMs2XO1V7Y',
    incomeStatementInstructions: 'https://drive.google.com/file/d/1Y8B4zPLb_DTkV2GZnlGztm-HMfA3OWYP/view?usp=sharing'
  };

  const SETTINGS_KEYS = {
    officeMap: 'email_link_office_map',
    ourTeam: 'email_link_our_team',
    factFind: 'email_link_fact_find',
    myGov: 'email_link_mygov',
    myGovVideo: 'email_link_mygov_video',
    incomeStatementInstructions: 'email_link_income_instructions'
  };

  let EMAIL_LINKS = { ...DEFAULT_EMAIL_LINKS };
  let userSignature = '';
  let emailSettingsLoaded = false;
  let currentUserProfile = null;
  let generatedSignatureHtml = '';

  function loadEmailLinksFromSettings() {
    google.script.run.withSuccessHandler(function(settings) {
      if (settings) {
        Object.keys(SETTINGS_KEYS).forEach(key => {
          const settingKey = SETTINGS_KEYS[key];
          if (settings[settingKey]) {
            EMAIL_LINKS[key] = settings[settingKey];
          }
        });
        emailSettingsLoaded = true;
        console.log('Email templates loaded:', Object.keys(settings).length);
      }
    }).getAllSettings();
  }

  function loadUserSignature() {
    google.script.run.withSuccessHandler(function(result) {
      if (result) {
        currentUserProfile = result;
        if (result.signature) {
          userSignature = result.signature;
        }
      }
    }).getUserSignature();
  }

  function openEmailSettings() {
    const officeMap = document.getElementById('settingOfficeMap');
    const ourTeam = document.getElementById('settingOurTeam');
    const factFind = document.getElementById('settingFactFind');
    const myGov = document.getElementById('settingMyGov');
    const myGovVideo = document.getElementById('settingMyGovVideo');
    const incomeInstructions = document.getElementById('settingIncomeInstructions');
    const signature = document.getElementById('settingSignature');
    const previewContainer = document.getElementById('signaturePreviewContainer');

    if (officeMap) officeMap.value = EMAIL_LINKS.officeMap || '';
    if (ourTeam) ourTeam.value = EMAIL_LINKS.ourTeam || '';
    if (factFind) factFind.value = EMAIL_LINKS.factFind || '';
    if (myGov) myGov.value = EMAIL_LINKS.myGov || '';
    if (myGovVideo) myGovVideo.value = EMAIL_LINKS.myGovVideo || '';
    if (incomeInstructions) incomeInstructions.value = EMAIL_LINKS.incomeStatementInstructions || '';
    if (signature) signature.value = userSignature || '';

    if (previewContainer) {
      if (userSignature) {
        previewContainer.innerHTML = userSignature;
        generatedSignatureHtml = userSignature;
      } else {
        previewContainer.innerHTML = '<span style="color:#999; font-style:italic;">No signature set. Click "Regenerate" to create one.</span>';
      }
    }

    if (currentUserProfile) {
      updateSignatureUserInfo();
    } else {
      google.script.run.withSuccessHandler(function(result) {
        if (result) {
          currentUserProfile = result;
          updateSignatureUserInfo();
        }
      }).getUserSignature();
    }

    openModal('emailSettingsModal');
  }

  function openGlobalSettings() {
    openEmailSettings();
  }

  function closeEmailSettings() {
    closeModal('emailSettingsModal');
  }

  function saveEmailSettings() {
    const newLinks = {
      officeMap: document.getElementById('settingOfficeMap').value,
      ourTeam: document.getElementById('settingOurTeam').value,
      factFind: document.getElementById('settingFactFind').value,
      myGov: document.getElementById('settingMyGov').value,
      myGovVideo: document.getElementById('settingMyGovVideo').value,
      incomeStatementInstructions: document.getElementById('settingIncomeInstructions').value
    };

    Object.assign(EMAIL_LINKS, newLinks);

    let saveCount = 0;
    let savedCount = 0;
    Object.keys(SETTINGS_KEYS).forEach(key => {
      const settingKey = SETTINGS_KEYS[key];
      saveCount++;
      google.script.run.withSuccessHandler(function() {
        savedCount++;
        if (savedCount === saveCount) {
          checkSignatureAndClose();
        }
      }).withFailureHandler(function(err) {
        savedCount++;
        console.error('Failed to save setting:', settingKey, err);
        if (savedCount === saveCount) {
          checkSignatureAndClose();
        }
      }).updateSetting(settingKey, newLinks[key]);
    });

    function checkSignatureAndClose() {
      const newSignature = document.getElementById('settingSignature').value;
      if (newSignature !== userSignature) {
        google.script.run.withSuccessHandler(function() {
          userSignature = newSignature;
          if (typeof updateEmailPreview === 'function') updateEmailPreview();
          showAlert('Saved', 'Settings and signature updated for the whole team', 'success');
        }).withFailureHandler(function(err) {
          showAlert('Error', 'Settings saved, but signature failed: ' + err, 'error');
        }).updateUserSignature(newSignature);
      } else {
        showAlert('Saved', 'Email template links updated for the whole team', 'success');
      }
      closeEmailSettings();
      if (typeof updateEmailPreview === 'function') updateEmailPreview();
    }
  }

  function updateSignatureUserInfo() {
    const nameEl = document.getElementById('sigGenName');
    const titleEl = document.getElementById('sigGenTitle');
    if (nameEl && currentUserProfile) {
      nameEl.innerText = currentUserProfile.name || 'Unknown';
    }
    if (titleEl && currentUserProfile) {
      titleEl.innerText = currentUserProfile.title || '';
    }
  }

  function regenerateSignature() {
    if (currentUserProfile) {
      generateSignaturePreview();
      const textarea = document.getElementById('settingSignature');
      if (textarea) {
        textarea.value = generatedSignatureHtml;
      }
      showAlert('Regenerated', 'Signature updated. Click Save to store it.', 'success');
    } else {
      google.script.run.withSuccessHandler(function(result) {
        if (result) {
          currentUserProfile = result;
          updateSignatureUserInfo();
          generateSignaturePreview();
          const textarea = document.getElementById('settingSignature');
          if (textarea) {
            textarea.value = generatedSignatureHtml;
          }
          showAlert('Regenerated', 'Signature updated. Click Save to store it.', 'success');
        }
      }).getUserSignature();
    }
  }

  function generateSignaturePreview() {
    if (!currentUserProfile) return;

    const name = currentUserProfile.name || 'Your Name';
    const title = currentUserProfile.title || 'Your Title';

    const signatureHtml = `<table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; font-size: 10pt; color: #333333;">
    <tbody>
        <tr>
            <td style="padding-bottom: 10px; line-height: 1.5;">
                Best wishes,<br><br>
                <strong style="font-size: 11pt;">${name}</strong><br>
                <strong>${title}</strong><br>
                <strong>Stellaris Finance Broking</strong>
            </td>
        </tr>
        <tr>
            <td style="line-height: 1.5;">
                Phone: 0488 839 212<br>
                Office: Suite 18 / 56 Creaney Drive, Kingsley WA 6026<br>
                Website: <a href="https://www.stellaris.loans" target="_blank" style="color: #1155cc; text-decoration: none;">www.stellaris.loans</a><br>
                Book an Appointment with Tim: <a href="https://calendly.com/tim-kerin" target="_blank" style="color: #1155cc; text-decoration: none;">calendly.com/tim-kerin</a>
            </td>
        </tr>
        <tr>
            <td style="padding-top: 15px;">
                <img 
                    src="https://img1.wsimg.com/isteam/ip/2c5f94ee-4964-4e9b-9b9c-a55121f8611b/WEB_Stellaris_Email%20Signature_Midnight.png" 
                    alt="Stellaris Finance Broking Email Signature Graphic" 
                    width="320" 
                    height="104" 
                    style="display: block; border: 0; width: 320px; height: 104px;">
            </td>
        </tr>
        <tr>
            <td style="padding-top: 20px; font-size: 9pt; font-style: italic; color: #888888; line-height: 1.4;">
                Credit Representative 379175 is authorised under Australian Credit Licence 389328
                <br><br>
                <strong>Confidentiality Notice</strong><br>
                This email and its attachments are confidential and intended for the recipient only. If you are not the intended recipient, please notify us and delete this message. Unauthorized use, dissemination, or copying is prohibited. The views expressed are those of the sender unless stated otherwise. We do not guarantee that attachments are free from viruses; the user assumes all responsibility for any resulting damage. We value your privacy. Your information may be used to provide financial services and may be shared with third parties as required by law.
                <br><br>
                <strong>Important Notice</strong><br>
                We will never ask you to transfer money or make payments via email. If you receive any such requests, please do not respond and contact us directly before taking any action. Your security is our priority. More information on identifying and protecting yourself from phishing attacks generally can be found on the government's ScamWatch website, www.scamwatch.gov.au.
            </td>
        </tr>
    </tbody>
</table>`;

    generatedSignatureHtml = signatureHtml;
    const previewContainer = document.getElementById('signaturePreviewContainer');
    if (previewContainer) {
      previewContainer.innerHTML = signatureHtml;
    }
  }

  async function copySignatureForGmail() {
    const previewEl = document.getElementById('signaturePreviewContainer');
    if (!previewEl || !generatedSignatureHtml) {
      showAlert('No Signature', 'Generate a signature first before copying.', 'error');
      return;
    }

    const gmailInstructions = `Signature has been copied to your clipboard. To update in Gmail:

1. Settings cog (top right of Gmail) then "See all settings"
2. Scroll down to signature, select "Stellaris Signature"
3. Click in the field to the right, remove everything from it and then Ctrl-V the new signature.
4. Scroll to the bottom and click Save Changes and you're done.`;

    try {
      if (navigator.clipboard && navigator.clipboard.write) {
        const htmlBlob = new Blob([generatedSignatureHtml], { type: 'text/html' });
        const textBlob = new Blob([previewEl.innerText], { type: 'text/plain' });
        const clipboardItem = new ClipboardItem({
          'text/html': htmlBlob,
          'text/plain': textBlob
        });
        await navigator.clipboard.write([clipboardItem]);
        showAlert('Copied for Gmail', gmailInstructions, 'success');
      } else {
        const range = document.createRange();
        range.selectNodeContents(previewEl);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        document.execCommand('copy');
        selection.removeAllRanges();
        showAlert('Copied for Gmail', gmailInstructions, 'success');
      }
    } catch (err) {
      const range = document.createRange();
      range.selectNodeContents(previewEl);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand('copy');
      selection.removeAllRanges();
      showAlert('Copied for Gmail', gmailInstructions, 'success');
    }
  }

  async function copySignatureForMercury() {
    const previewEl = document.getElementById('signaturePreviewContainer');
    if (!previewEl || !generatedSignatureHtml) {
      showAlert('No Signature', 'Generate a signature first before copying.', 'error');
      return;
    }

    const mercuryInstructions = `Signature has been copied to your clipboard. To paste in Mercury, log into Mercury then:

1. Admin tile
2. Email Profiles tab
3. On the left, select the Profile you want to update the signature for
4. On the right, Edit Details tab
5. In the Email Signature pane, click the three dots in the top right (More misc) then < > (Code view)
6. Delete all that code
7. Ctrl-V the code you just copied here in Integrity
8. Click < > (Code View) again to turn it off and make sure the rendered sig looks right, then click the Preview tab to be sure. If it's good, you're done.`;

    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(generatedSignatureHtml);
        showAlert('Copied for Mercury', mercuryInstructions, 'success');
      } else {
        throw new Error('Clipboard API not available');
      }
    } catch (err) {
      const textarea = document.createElement('textarea');
      textarea.value = generatedSignatureHtml;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      showAlert('Copied for Mercury', mercuryInstructions, 'success');
    }
  }

  function useGeneratedSignature() {
    const textarea = document.getElementById('settingSignature');
    if (textarea) {
      textarea.value = generatedSignatureHtml;
    }

    const previewContainer = document.getElementById('signaturePreviewContainer');
    if (previewContainer) {
      previewContainer.innerHTML = generatedSignatureHtml;
    }

    showAlert('Applied!', 'Signature added. Click Save to store it.', 'success');
  }

  loadEmailLinksFromSettings();
  loadUserSignature();

  window.DEFAULT_EMAIL_LINKS = DEFAULT_EMAIL_LINKS;
  window.SETTINGS_KEYS = SETTINGS_KEYS;
  window.EMAIL_LINKS = EMAIL_LINKS;
  
  Object.defineProperty(window, 'userSignature', {
    get: function() { return userSignature; },
    set: function(val) { userSignature = val; }
  });
  Object.defineProperty(window, 'emailSettingsLoaded', {
    get: function() { return emailSettingsLoaded; },
    set: function(val) { emailSettingsLoaded = val; }
  });
  Object.defineProperty(window, 'currentUserProfile', {
    get: function() { return currentUserProfile; },
    set: function(val) { currentUserProfile = val; }
  });
  Object.defineProperty(window, 'generatedSignatureHtml', {
    get: function() { return generatedSignatureHtml; },
    set: function(val) { generatedSignatureHtml = val; }
  });

  window.loadEmailLinksFromSettings = loadEmailLinksFromSettings;
  window.loadUserSignature = loadUserSignature;
  window.openEmailSettings = openEmailSettings;
  window.closeEmailSettings = closeEmailSettings;
  window.saveEmailSettings = saveEmailSettings;
  window.openGlobalSettings = openGlobalSettings;
  window.updateSignatureUserInfo = updateSignatureUserInfo;
  window.regenerateSignature = regenerateSignature;
  window.generateSignaturePreview = generateSignaturePreview;
  window.copySignatureForGmail = copySignatureForGmail;
  window.copySignatureForMercury = copySignatureForMercury;
  window.useGeneratedSignature = useGeneratedSignature;

})();
