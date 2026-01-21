/**
 * email.js - Email Composer and Template Management Module
 * 
 * Handles email composition with Quill WYSIWYG editor, template management,
 * conditional template parsing, and Gmail integration.
 * 
 * Dependencies: shared-state.js, shared-utils.js, modal-utils.js, settings.js
 * 
 * Functions exposed to window:
 * - openEmailComposer, closeEmailComposer, sendEmail, updateEmailPreview
 * - openEmailComposerFromPanel, openInGmail
 * - loadEmailTemplates, openTemplateList, closeTemplateList
 * - openTemplateEditor, closeTemplateEditor, saveTemplate
 * - insertVariable, insertConditionBlock, updateConditionOptions
 * - openCurrentTemplateEditor, createNewTemplate, seedDefaultTemplate
 * - updateTemplatePreviewFromControls
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // Email context and Quill editor
  let currentEmailContext = null;
  let emailQuill = null;
  
  // Airtable templates
  let airtableTemplates = [];
  let templatesLoaded = false;
  let currentEditingTemplate = null;
  
  // Template editor state
  let templateEditorQuill = null;
  let templateSubjectQuill = null;
  let activeTemplateEditor = 'body';
  let templatePreviewContext = null;
  let previewUpdateTimeout = null;
  let isHighlighting = false;
  
  // Hardcoded email template (fallback)
  const EMAIL_TEMPLATE = {
    subject: {
      Office: 'Confirmation of Appointment - {{appointmentTime}}',
      Phone: 'Confirmation of Phone Appointment - {{appointmentTime}}',
      Video: 'Confirmation of Google Meet Appointment - {{appointmentTime}}'
    },
    
    opening: {
      Office: "I'm writing to confirm your appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) at our office - Kingsley Professional Centre, 18 / 56 Creaney Drive, Kingsley ({{officeMapLink}}).",
      Phone: "I'm writing to confirm your phone appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today). {{brokerFirst}} will call you on {{phoneNumber}}.",
      Video: "I'm writing to confirm your video call appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) using Google Meet URL: {{meetUrl}}. If you have any trouble logging in, please call or text our team on 0488 839 212."
    },
    
    meetLine: {
      New: "Click {{ourTeamLink}} to meet {{brokerFirst}} and the rest of the Team at Stellaris Finance Broking. We will be supporting you each step of the way!",
      Repeat: "Click {{ourTeamLink}} to get reacquainted with the Team at Stellaris Finance Broking. We will be supporting you each step of the way!"
    },
    
    preparation: {
      Shae: `In preparation for your appointment, please email me the following information:

{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)
- Complete with as much detail as possible
- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3

Income
a) PAYG Income – Your latest two consecutive payslips and your 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})
b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment

I work part time – please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.`,
      Team: `In preparation for your appointment, please email Shae (shae@stellaris.loans) the following information:

{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)
- Complete with as much detail as possible
- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3

Income
a) PAYG Income – Your latest two consecutive payslips and 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})
b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment

Please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.`,
      OpenBanking: `You will soon receive invitations to share your information with us via Frollo's Open Banking and Connective's Client Centre. These two systems streamline the collection of your key financial and personal data, including your contact details, employment history, savings and liabilities, to give us a complete picture of your situation.{{prefillNote}}`
    },
    
    prefillNote: {
      New: '',
      Repeat: ' I have prefilled as much as I can using the information we previously received from you.'
    },
    
    closing: {
      New: `Do not hesitate to contact our team on 0488 839 212 if you have any questions.

We look forward to working with you!

Best wishes,
{{sender}}`,
      Repeat: `Do not hesitate to contact our team on 0488 839 212 if you have any questions.

We look forward to working with you again!

Best wishes,
{{sender}}`
    }
  };
  
  // Template variable definitions for picker
  const TEMPLATE_VARIABLES = {
    'Client Info': [
      { name: 'greeting', description: 'Client first name or preferred greeting' },
      { name: 'clientType', description: '"New" or "Repeat"' }
    ],
    'Broker Info': [
      { name: 'broker', description: 'Full broker name' },
      { name: 'brokerFirst', description: 'Broker first name' },
      { name: 'brokerIntro', description: '"our Mortgage Broker [Name]" for new clients, just first name for repeat' },
      { name: 'sender', description: 'Current user (email sender) name' }
    ],
    'Appointment Details': [
      { name: 'appointmentType', description: '"Office", "Phone", or "Video"' },
      { name: 'appointmentTime', description: 'Formatted appointment date/time' },
      { name: 'daysUntil', description: 'Number of days until appointment' },
      { name: 'phoneNumber', description: 'Client phone number' },
      { name: 'meetUrl', description: 'Google Meet URL for video calls' }
    ],
    'Preparation': [
      { name: 'prepHandler', description: '"Shae", "Team", or "OpenBanking"' },
      { name: 'prefillNote', description: 'Auto-fill note based on client type' }
    ],
    'Links': [
      { name: 'officeMapLink', description: 'Office map link (clickable)' },
      { name: 'ourTeamLink', description: 'Team page link (clickable)' },
      { name: 'factFindLink', description: 'Fact Find document link' },
      { name: 'myGovLink', description: 'myGov link' },
      { name: 'myGovVideoLink', description: 'myGov help video link' },
      { name: 'incomeInstructionsLink', description: 'Income statement instructions link' }
    ]
  };
  
  // Condition variables for {{if}} blocks
  const CONDITION_VARIABLES = [
    { name: 'appointmentType', options: ['Office', 'Phone', 'Video'] },
    { name: 'clientType', options: ['New', 'Repeat'] },
    { name: 'prepHandler', options: ['Shae', 'Team', 'OpenBanking'] }
  ];
  
  // Sample data for preview when no opportunity is selected
  const SAMPLE_PREVIEW_DATA = {
    greeting: 'Sarah',
    clientType: 'New',
    broker: 'Michael Thompson',
    brokerFirst: 'Michael',
    brokerIntro: 'our Mortgage Broker Michael Thompson',
    sender: 'Shae',
    appointmentType: 'Office',
    appointmentTime: 'Tuesday 15th January at 10:00am',
    daysUntil: '3',
    phoneNumber: '0412 345 678',
    meetUrl: 'https://meet.google.com/abc-defg-hij',
    prepHandler: 'Team',
    prefillNote: '',
    officeMapLink: '<a href="#" style="color:#0066CC;">Office</a>',
    ourTeamLink: '<a href="#" style="color:#0066CC;">here</a>',
    factFindLink: '<a href="#" style="color:#0066CC;">Fact Find</a>',
    myGovLink: '<a href="#" style="color:#0066CC;">myGov</a>',
    myGovVideoLink: '<a href="#" style="color:#0066CC;">myGov Video</a>',
    incomeInstructionsLink: '<a href="#" style="color:#0066CC;">Income Instructions</a>'
  };

  // ============================================================
  // Core Email Composer Functions
  // ============================================================
  
  function openEmailComposer(opportunityData, contactData) {
    const emails = [];
    if (contactData.EmailAddress1) emails.push(contactData.EmailAddress1);
    
    if (opportunityData._applicantEmails && Array.isArray(opportunityData._applicantEmails)) {
      opportunityData._applicantEmails.forEach(email => {
        if (email && !emails.includes(email)) emails.push(email);
      });
    }
    
    const apptData = opportunityData._appointmentData || {};
    
    currentEmailContext = {
      opportunity: opportunityData,
      contact: contactData,
      greeting: contactData.PreferredName || contactData.FirstName || 'there',
      broker: opportunityData['Taco: Broker'] || 'our Mortgage Broker',
      brokerFirst: (opportunityData['Taco: Broker'] || '').split(' ')[0] || 'the broker',
      appointmentTime: formatAppointmentTime(apptData.appointmentTime || opportunityData['Taco: Appointment Time']),
      phoneNumber: apptData.phoneNumber || opportunityData['Taco: Appt Phone Number'] || '[phone number]',
      meetUrl: apptData.meetUrl || opportunityData['Taco: Appt Meet URL'] || '[Google Meet URL]',
      emails: emails,
      sender: 'Shae'
    };
    
    const apptType = apptData.typeOfAppointment || opportunityData['Taco: Type of Appointment'] || 'Phone';
    document.getElementById('emailApptType').value = apptType;
    
    const isNew = opportunityData['Taco: New or Existing Client'] === 'New Client';
    document.getElementById('emailClientType').value = isNew ? 'New' : 'Repeat';
    
    document.getElementById('emailPrepHandler').value = 'Shae';
    document.getElementById('emailTo').value = currentEmailContext.emails.join(', ');
    
    const modal = document.getElementById('emailComposer');
    modal.classList.add('visible');
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        modal.classList.add('showing');
      });
    });
    
    // Initialize Quill editor if not already done
    if (!emailQuill && typeof Quill !== 'undefined') {
      emailQuill = new Quill('#emailPreviewBody', {
        modules: {
          toolbar: '#emailQuillToolbar'
        },
        theme: 'snow'
      });
    }
    
    updateEmailPreview();
  }
  
  function closeEmailComposer() {
    const modal = document.getElementById('emailComposer');
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.classList.remove('visible');
      currentEmailContext = null;
    }, 250);
  }
  
  function sendEmail() {
    if (!currentEmailContext) return;
    
    const to = document.getElementById('emailTo').value;
    const subject = document.getElementById('emailSubject').value;
    const body = emailQuill ? emailQuill.root.innerHTML : document.getElementById('emailPreviewBody').innerHTML;
    
    const sendBtn = document.getElementById('emailSendBtn');
    sendBtn.innerText = 'Sending...';
    sendBtn.disabled = true;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.success) {
        showAlert('Success', 'Email sent successfully!', 'success');
        closeEmailComposer();
      } else {
        showAlert('Error', result?.error || 'Failed to send email. Gmail API integration required.', 'error');
      }
      sendBtn.innerText = 'Send';
      sendBtn.disabled = false;
    }).withFailureHandler(function(err) {
      showAlert('Error', 'Failed to send email: ' + (err.message || 'Gmail API integration required.'), 'error');
      sendBtn.innerText = 'Send';
      sendBtn.disabled = false;
    }).sendEmail(to, subject, body);
  }
  
  function updateEmailPreview() {
    if (!currentEmailContext) return;
    
    const apptType = document.getElementById('emailApptType').value;
    const clientType = document.getElementById('emailClientType').value;
    const prepHandler = document.getElementById('emailPrepHandler').value;
    
    // Access EMAIL_LINKS from settings.js module
    const EMAIL_LINKS = window.EMAIL_LINKS || {};
    const userSignature = window.userSignature || '';
    
    const brokerIntro = clientType === 'New' 
      ? `our Mortgage Broker ${currentEmailContext.broker}`
      : currentEmailContext.brokerFirst;
    
    const variables = {
      greeting: currentEmailContext.greeting,
      broker: currentEmailContext.broker,
      brokerFirst: currentEmailContext.brokerFirst,
      brokerIntro: brokerIntro,
      appointmentTime: currentEmailContext.appointmentTime,
      appointmentType: apptType,
      clientType: clientType,
      prepHandler: prepHandler,
      daysUntil: calculateDaysUntil(currentEmailContext.appointmentTime),
      phoneNumber: currentEmailContext.phoneNumber,
      meetUrl: currentEmailContext.meetUrl,
      sender: currentEmailContext.sender,
      prefillNote: EMAIL_TEMPLATE.prefillNote[clientType],
      officeMapLink: `<a href="${EMAIL_LINKS.officeMap || '#'}" target="_blank" style="color:#0066CC;">Office</a>`,
      ourTeamLink: `<a href="${EMAIL_LINKS.ourTeam || '#'}" target="_blank" style="color:#0066CC;">here</a>`,
      factFindLink: `<a href="${EMAIL_LINKS.factFind || '#'}" target="_blank" style="color:#0066CC;">Fact Find</a>`,
      myGovLink: `<a href="${EMAIL_LINKS.myGov || '#'}" target="_blank" style="color:#0066CC;">myGov</a>`,
      myGovVideoLink: `<a href="${EMAIL_LINKS.myGovVideo || '#'}" target="_blank" style="color:#0066CC;">this video</a>`,
      incomeInstructionsLink: `<a href="${EMAIL_LINKS.incomeStatementInstructions || '#'}" target="_blank" style="color:#0066CC;">click here</a>`
    };
    
    const confirmationTemplate = getTemplateByType('Confirmation') || getTemplateByName('Appointment Confirmation');
    
    if (confirmationTemplate && confirmationTemplate.body) {
      const subject = parseConditionalTemplate(confirmationTemplate.subject || EMAIL_TEMPLATE.subject[apptType], variables);
      document.getElementById('emailSubject').value = subject;
      
      let body = parseConditionalTemplate(confirmationTemplate.body, variables);
      
      if (userSignature) {
        body += '<br><br>' + userSignature.replace(/\n/g, '<br>');
      }
      
      if (emailQuill) {
        emailQuill.clipboard.dangerouslyPasteHTML(body);
      } else {
        document.getElementById('emailPreviewBody').innerHTML = body;
      }
    } else {
      const subject = replaceVariables(EMAIL_TEMPLATE.subject[apptType], variables);
      document.getElementById('emailSubject').value = subject;
      
      let body = `Hi ${variables.greeting},<br><br>`;
      body += replaceVariables(EMAIL_TEMPLATE.opening[apptType], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.meetLine[clientType], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.preparation[prepHandler], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.closing[clientType], variables);
      
      if (userSignature) {
        body += '<br><br>' + userSignature.replace(/\n/g, '<br>');
      }
      
      if (emailQuill) {
        emailQuill.clipboard.dangerouslyPasteHTML(body);
      } else {
        document.getElementById('emailPreviewBody').innerHTML = body;
      }
    }
  }
  
  // ============================================================
  // Template Parsing Functions
  // ============================================================
  
  function replaceVariables(template, variables) {
    if (!template) return '';
    return template.replace(/\{\{(\w+)\}\}/g, (match, key) => variables[key] || match);
  }
  
  function parseConditionalTemplate(template, context) {
    if (!template) return '';
    let result = processConditionals(template, context);
    result = replaceVariables(result, context);
    return result;
  }
  
  function processConditionals(template, context) {
    const conditionalPattern = /\{\{if\s+(\w+)\s*=\s*(\w+)\}\}([\s\S]*?)\{\{endif\}\}/gi;
    
    return template.replace(conditionalPattern, function(match, varName, varValue, innerContent) {
      const branches = parseConditionalBranches(innerContent);
      const contextValue = context[varName];
      
      for (const branch of branches) {
        if (branch.type === 'if' || branch.type === 'elseif') {
          if (branch.value === contextValue) {
            return processConditionals(branch.content, context);
          }
        } else if (branch.type === 'else') {
          return processConditionals(branch.content, context);
        }
      }
      
      if (varValue === contextValue) {
        const firstBranchContent = innerContent.split(/\{\{(?:elseif|else)/i)[0];
        return processConditionals(firstBranchContent, context);
      }
      
      return '';
    });
  }
  
  function parseConditionalBranches(content) {
    const branches = [];
    const parts = content.split(/(\{\{elseif\s+\w+\s*=\s*\w+\}\}|\{\{else\}\})/gi);
    
    let currentBranch = { type: 'if', content: parts[0] || '', value: null };
    
    for (let i = 1; i < parts.length; i++) {
      const part = parts[i];
      
      if (/\{\{elseif\s+(\w+)\s*=\s*(\w+)\}\}/i.test(part)) {
        branches.push(currentBranch);
        const match = part.match(/\{\{elseif\s+(\w+)\s*=\s*(\w+)\}\}/i);
        currentBranch = { type: 'elseif', varName: match[1], value: match[2], content: '' };
      } else if (/\{\{else\}\}/i.test(part)) {
        branches.push(currentBranch);
        currentBranch = { type: 'else', content: '' };
      } else {
        currentBranch.content += part;
      }
    }
    
    branches.push(currentBranch);
    return branches;
  }
  
  // ============================================================
  // Template Management Functions
  // ============================================================
  
  function loadEmailTemplates() {
    google.script.run.withSuccessHandler(function(templates) {
      airtableTemplates = templates || [];
      templatesLoaded = true;
      console.log('Email templates loaded:', airtableTemplates.length);
    }).withFailureHandler(function(err) {
      console.error('Failed to load email templates:', err);
      airtableTemplates = [];
      templatesLoaded = true;
    }).getEmailTemplates();
  }
  
  function getTemplateByName(name) {
    return airtableTemplates.find(t => t.name.toLowerCase() === name.toLowerCase());
  }
  
  function getTemplateByType(type) {
    return airtableTemplates.find(t => t.type === type);
  }
  
  // ============================================================
  // Template List (Central Hub)
  // ============================================================
  
  function openTemplateList() {
    const modal = document.getElementById('templateListModal');
    if (!modal) return;
    
    refreshTemplateList();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  }
  
  function closeTemplateList() {
    const modal = document.getElementById('templateListModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => { modal.style.display = 'none'; }, 250);
    }
  }
  
  function refreshTemplateList() {
    const container = document.getElementById('templateListContainer');
    if (!container) return;
    
    if (airtableTemplates.length === 0) {
      container.innerHTML = `<div style="text-align:center; padding:30px;">
        <div style="color:#888; font-style:italic; margin-bottom:15px;">No templates yet.</div>
        <button type="button" onclick="seedDefaultTemplate()" style="padding:10px 20px; background:var(--color-star); color:white; border:none; border-radius:6px; cursor:pointer; font-size:13px; font-weight:500;">Load Default Confirmation Template</button>
        <div style="font-size:11px; color:#999; margin-top:10px;">This will create the standard Appointment Confirmation template with all variations.</div>
      </div>`;
      return;
    }
    
    let html = '';
    airtableTemplates.forEach(t => {
      html += `<div class="template-list-item">`;
      html += `<div><span class="template-list-name">${t.name}</span></div>`;
      html += `<div class="template-list-actions">`;
      html += `<span class="template-list-type">${t.type || 'General'}</span>`;
      html += `<button class="template-list-edit" onclick="closeTemplateList(); openTemplateEditor('${t.id}')">Edit</button>`;
      html += `</div>`;
      html += `</div>`;
    });
    container.innerHTML = html;
  }
  
  function createNewTemplate() {
    closeTemplateList();
    openTemplateEditor(null);
  }
  
  function seedDefaultTemplate() {
    const defaultSubject = `{{if appointmentType=Office}}Confirmation of Appointment - {{appointmentTime}}{{elseif appointmentType=Phone}}Confirmation of Phone Appointment - {{appointmentTime}}{{elseif appointmentType=Video}}Confirmation of Google Meet Appointment - {{appointmentTime}}{{endif}}`;
    
    const defaultBody = `Hi {{greeting}},<br><br>{{if appointmentType=Office}}I'm writing to confirm your appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) at our office - Kingsley Professional Centre, 18 / 56 Creaney Drive, Kingsley ({{officeMapLink}}).{{elseif appointmentType=Phone}}I'm writing to confirm your phone appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today). {{brokerFirst}} will call you on {{phoneNumber}}.{{elseif appointmentType=Video}}I'm writing to confirm your video call appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) using Google Meet URL: {{meetUrl}}. If you have any trouble logging in, please call or text our team on 0488 839 212.{{endif}}<br><br>{{if clientType=New}}Click {{ourTeamLink}} to meet {{brokerFirst}} and the rest of the Team at Stellaris Finance Broking. We will be supporting you each step of the way!{{elseif clientType=Repeat}}Click {{ourTeamLink}} to get reacquainted with the Team at Stellaris Finance Broking. We will be supporting you each step of the way!{{endif}}<br><br>{{if prepHandler=Shae}}In preparation for your appointment, please email me the following information:<br><br>{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)<br>- Complete with as much detail as possible<br>- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3<br><br>Income<br>a) PAYG Income – Your latest two consecutive payslips and your 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})<br>b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment<br><br>I work part time – please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.{{elseif prepHandler=Team}}In preparation for your appointment, please email Shae (shae@stellaris.loans) the following information:<br><br>{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)<br>- Complete with as much detail as possible<br>- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3<br><br>Income<br>a) PAYG Income – Your latest two consecutive payslips and 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})<br>b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment<br><br>Please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.{{elseif prepHandler=OpenBanking}}You will soon receive invitations to share your information with us via Frollo's Open Banking and Connective's Client Centre.{{prefillNote}}{{endif}}<br><br>{{if clientType=New}}Do not hesitate to contact our team on 0488 839 212 if you have any questions.<br><br>We look forward to working with you!<br><br>Best wishes,<br>{{sender}}{{elseif clientType=Repeat}}Do not hesitate to contact our team on 0488 839 212 if you have any questions.<br><br>We look forward to working with you again!<br><br>Best wishes,<br>{{sender}}{{endif}}`;
    
    showAlert('Creating...', 'Creating default template...', 'success');
    
    google.script.run.withSuccessHandler(function(result) {
      if (result) {
        showAlert('Success', 'Default template created! Refresh to see it.', 'success');
        loadEmailTemplates();
        refreshTemplateList();
      } else {
        showAlert('Error', 'Failed to create template', 'error');
      }
    }).withFailureHandler(function(err) {
      showAlert('Error', 'Failed to create template: ' + (err.message || 'Unknown error'), 'error');
    }).createEmailTemplate({
      name: 'Appointment Confirmation',
      type: 'Confirmation',
      subject: defaultSubject,
      body: defaultBody,
      description: 'Standard appointment confirmation email with variations for appointment type, client type, and preparation method.',
      active: true
    });
  }
  
  // ============================================================
  // Template Editor Functions
  // ============================================================
  
  function openTemplateEditor(templateId) {
    console.log('openTemplateEditor called with:', templateId);
    const modal = document.getElementById('templateEditorModal');
    if (!modal) {
      console.error('Template editor modal not found');
      return;
    }
    
    try {
      if (!templateSubjectQuill && typeof Quill !== 'undefined') {
        templateSubjectQuill = new Quill('#templateEditorSubject', {
          modules: { toolbar: false },
          theme: 'snow',
          placeholder: 'Email subject with {{variables}}'
        });
        templateSubjectQuill.on('text-change', () => {
          highlightVariables(templateSubjectQuill);
          schedulePreviewUpdate();
        });
        templateSubjectQuill.root.addEventListener('focus', () => { activeTemplateEditor = 'subject'; });
      }
      
      if (!templateEditorQuill && typeof Quill !== 'undefined') {
        templateEditorQuill = new Quill('#templateEditorBody', {
          modules: { toolbar: '#templateEditorToolbar' },
          theme: 'snow'
        });
        templateEditorQuill.on('text-change', () => {
          highlightVariables(templateEditorQuill);
          schedulePreviewUpdate();
        });
        templateEditorQuill.root.addEventListener('focus', () => { activeTemplateEditor = 'body'; });
      }
    } catch (err) {
      console.error('Error initializing Quill editors:', err);
    }
    
    populateVariablePicker();
    
    const conditionVarSelect = document.getElementById('conditionVariable');
    if (conditionVarSelect) {
      conditionVarSelect.onchange = function() {
        updateConditionOptions(this.value);
      };
    }
    
    if (templateId) {
      const template = airtableTemplates.find(t => t.id === templateId);
      if (template) {
        currentEditingTemplate = template;
        document.getElementById('templateEditorTitle').innerText = 'Edit Template';
        document.getElementById('templateEditorName').value = template.name;
        if (templateSubjectQuill) {
          templateSubjectQuill.setText(template.subject || '');
          setTimeout(() => highlightVariables(templateSubjectQuill), 50);
        }
        if (templateEditorQuill) {
          templateEditorQuill.clipboard.dangerouslyPasteHTML(template.body);
          setTimeout(() => highlightVariables(templateEditorQuill), 50);
        }
      }
    } else {
      currentEditingTemplate = null;
      document.getElementById('templateEditorTitle').innerText = 'New Template';
      document.getElementById('templateEditorName').value = '';
      if (templateSubjectQuill) templateSubjectQuill.setText('');
      if (templateEditorQuill) templateEditorQuill.setText('');
    }
    
    const dataBadge = document.getElementById('previewDataSource');
    if (dataBadge) {
      if (templatePreviewContext) {
        dataBadge.textContent = 'Live Data';
        dataBadge.classList.add('live-data');
      } else {
        dataBadge.textContent = 'Sample Data';
        dataBadge.classList.remove('live-data');
      }
    }
    
    modal.style.display = 'flex';
    setTimeout(() => {
      modal.classList.add('showing');
      initializePreviewControls();
      setTimeout(renderTemplatePreview, 100);
    }, 10);
  }
  
  function closeTemplateEditor() {
    const modal = document.getElementById('templateEditorModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => { modal.style.display = 'none'; }, 250);
    }
    currentEditingTemplate = null;
    templatePreviewContext = null;
  }
  
  function saveTemplate() {
    const name = document.getElementById('templateEditorName').value.trim();
    const subject = templateSubjectQuill ? templateSubjectQuill.getText().trim() : '';
    const body = templateEditorQuill ? templateEditorQuill.root.innerHTML : '';
    
    if (!name) {
      showAlert('Error', 'Please enter a template name', 'error');
      return;
    }
    
    const btn = document.getElementById('templateEditorSaveBtn');
    btn.innerText = 'Saving...';
    btn.disabled = true;
    
    const fields = { name, subject, body };
    
    if (currentEditingTemplate) {
      google.script.run.withSuccessHandler(function(result) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        if (result) {
          showAlert('Success', 'Template saved successfully', 'success');
          loadEmailTemplates();
          closeTemplateEditor();
        } else {
          showAlert('Error', 'Failed to save template', 'error');
        }
      }).withFailureHandler(function(err) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        showAlert('Error', 'Failed to save template: ' + (err.message || 'Unknown error'), 'error');
      }).updateEmailTemplate(currentEditingTemplate.id, fields);
    } else {
      google.script.run.withSuccessHandler(function(result) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        if (result) {
          showAlert('Success', 'Template created successfully', 'success');
          loadEmailTemplates();
          closeTemplateEditor();
        } else {
          showAlert('Error', 'Failed to create template', 'error');
        }
      }).withFailureHandler(function(err) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        showAlert('Error', 'Failed to create template: ' + (err.message || 'Unknown error'), 'error');
      }).createEmailTemplate(fields);
    }
  }
  
  // ============================================================
  // Template Editor Helpers
  // ============================================================
  
  function highlightVariables(quillInstance) {
    if (!quillInstance || isHighlighting) return;
    isHighlighting = true;
    
    try {
      const text = quillInstance.getText();
      const regex = /\{\{[^}]+\}\}/g;
      let match;
      
      quillInstance.formatText(0, text.length, 'background', false);
      
      while ((match = regex.exec(text)) !== null) {
        quillInstance.formatText(match.index, match[0].length, 'background', '#D0DFE6');
      }
    } finally {
      isHighlighting = false;
    }
  }
  
  function populateVariablePicker() {
    const container = document.getElementById('variablePickerList');
    if (!container) return;
    
    let html = '';
    for (const [groupName, variables] of Object.entries(TEMPLATE_VARIABLES)) {
      html += `<div class="variable-group">`;
      html += `<div class="variable-group-title">${groupName}</div>`;
      for (const v of variables) {
        html += `<div class="variable-item" onclick="insertVariable('${v.name}')" title="${v.description}">`;
        html += `<span class="variable-item-name">{{${v.name}}}</span>`;
        html += `<span class="variable-item-desc">${v.description}</span>`;
        html += `</div>`;
      }
      html += `</div>`;
    }
    container.innerHTML = html;
  }
  
  function insertVariable(varName) {
    const quill = activeTemplateEditor === 'subject' ? templateSubjectQuill : templateEditorQuill;
    if (!quill) return;
    
    const range = quill.getSelection();
    const insertPos = range ? range.index : quill.getLength() - 1;
    quill.insertText(insertPos, `{{${varName}}}`);
    quill.setSelection(insertPos + varName.length + 4);
  }
  
  function updateConditionOptions(varName) {
    const container = document.getElementById('conditionOptionsContainer');
    const optionsDiv = document.getElementById('conditionOptions');
    
    if (!varName) {
      container.style.display = 'none';
      return;
    }
    
    const condVar = CONDITION_VARIABLES.find(v => v.name === varName);
    if (!condVar) {
      container.style.display = 'none';
      return;
    }
    
    let html = `<div class="condition-option-label">Select options to include:</div>`;
    condVar.options.forEach(opt => {
      html += `<div class="condition-option-row">`;
      html += `<input type="checkbox" id="condOpt_${opt}" value="${opt}" checked>`;
      html += `<span>${opt}</span>`;
      html += `</div>`;
    });
    
    optionsDiv.innerHTML = html;
    container.style.display = 'block';
  }
  
  function insertConditionBlock() {
    if (!templateEditorQuill) return;
    
    const varName = document.getElementById('conditionVariable').value;
    if (!varName) {
      showAlert('Error', 'Please select a variable first', 'error');
      return;
    }
    
    const condVar = CONDITION_VARIABLES.find(v => v.name === varName);
    if (!condVar) return;
    
    const selectedOptions = [];
    condVar.options.forEach(opt => {
      const checkbox = document.getElementById(`condOpt_${opt}`);
      if (checkbox && checkbox.checked) {
        selectedOptions.push(opt);
      }
    });
    
    if (selectedOptions.length === 0) {
      showAlert('Error', 'Please select at least one option', 'error');
      return;
    }
    
    let block = '';
    selectedOptions.forEach((opt, idx) => {
      if (idx === 0) {
        block += `{{if ${varName}=${opt}}}\n[Content for ${opt}]\n`;
      } else {
        block += `{{elseif ${varName}=${opt}}}\n[Content for ${opt}]\n`;
      }
    });
    block += `{{endif}}`;
    
    const range = templateEditorQuill.getSelection();
    const insertPos = range ? range.index : templateEditorQuill.getLength();
    templateEditorQuill.insertText(insertPos, block);
    
    document.getElementById('conditionVariable').value = '';
    document.getElementById('conditionOptionsContainer').style.display = 'none';
  }
  
  // ============================================================
  // Template Preview Functions
  // ============================================================
  
  function schedulePreviewUpdate() {
    if (previewUpdateTimeout) clearTimeout(previewUpdateTimeout);
    previewUpdateTimeout = setTimeout(renderTemplatePreview, 150);
  }
  
  function renderTemplatePreview() {
    if (!templateSubjectQuill || !templateEditorQuill) return;
    
    const subjectText = templateSubjectQuill.getText().trim();
    const bodyHtml = templateEditorQuill.root.innerHTML;
    
    const context = getPreviewContextWithOverrides();
    
    const renderedSubject = renderWithMissingIndicators(subjectText, context);
    const subjectEl = document.getElementById('templatePreviewSubject');
    if (subjectEl) subjectEl.innerHTML = renderedSubject;
    
    const renderedBody = renderBodyWithMissingIndicators(bodyHtml, context);
    const bodyEl = document.getElementById('templatePreviewBody');
    if (bodyEl) bodyEl.innerHTML = renderedBody;
  }
  
  function getPreviewContextWithOverrides() {
    const baseContext = templatePreviewContext || SAMPLE_PREVIEW_DATA;
    
    const apptTypeEl = document.getElementById('previewApptType');
    const clientTypeEl = document.getElementById('previewClientType');
    const prepHandlerEl = document.getElementById('previewPrepHandler');
    
    const appointmentType = apptTypeEl ? apptTypeEl.value : baseContext.appointmentType;
    const clientType = clientTypeEl ? clientTypeEl.value : baseContext.clientType;
    const prepHandler = prepHandlerEl ? prepHandlerEl.value : baseContext.prepHandler;
    
    const brokerIntro = clientType === 'New' 
      ? `our Mortgage Broker ${baseContext.broker}`
      : baseContext.brokerFirst;
    
    const prefillNote = clientType === 'Repeat' 
      ? ' I have prefilled as much as I can using the information we previously received from you.'
      : '';
    
    return {
      ...baseContext,
      appointmentType,
      clientType,
      prepHandler,
      brokerIntro,
      prefillNote
    };
  }
  
  function initializePreviewControls() {
    const context = templatePreviewContext || SAMPLE_PREVIEW_DATA;
    
    const apptTypeEl = document.getElementById('previewApptType');
    const clientTypeEl = document.getElementById('previewClientType');
    const prepHandlerEl = document.getElementById('previewPrepHandler');
    
    if (apptTypeEl) apptTypeEl.value = context.appointmentType || 'Office';
    if (clientTypeEl) clientTypeEl.value = context.clientType || 'New';
    if (prepHandlerEl) prepHandlerEl.value = context.prepHandler || 'Team';
  }
  
  function updateTemplatePreviewFromControls() {
    renderTemplatePreview();
  }
  
  function renderWithMissingIndicators(template, context) {
    if (!template) return '';
    
    let result = processConditionalsForPreview(template, context);
    
    result = result.replace(/\{\{(\w+)\}\}/g, (match, varName) => {
      const value = context[varName];
      if (value !== undefined && value !== null && value !== '') {
        return value;
      }
      return `<span class="preview-missing-var">[${varName} not set]</span>`;
    });
    
    return result;
  }
  
  function renderBodyWithMissingIndicators(html, context) {
    if (!html) return '';
    
    let result = processConditionalsForPreview(html, context);
    
    result = result.replace(/\{\{(\w+)\}\}/g, (match, varName) => {
      const value = context[varName];
      if (value !== undefined && value !== null && value !== '') {
        return value;
      }
      return `<span class="preview-missing-var">[${varName} not set]</span>`;
    });
    
    return result;
  }
  
  function processConditionalsForPreview(template, context) {
    if (!template) return '';
    
    const conditionalPattern = /\{\{if\s+(\w+)\s*=\s*(\w+)\}\}([\s\S]*?)\{\{endif\}\}/gi;
    
    return template.replace(conditionalPattern, function(match, varName, varValue, innerContent) {
      try {
        const contextValue = context[varName];
        const branches = parseConditionalBranches(innerContent);
        
        if (!branches || branches.length === 0) {
          return '';
        }
        
        branches[0].value = varValue;
        
        let matchedContent = '';
        let matchedCondition = '';
        
        for (const branch of branches) {
          if (branch.type === 'if' || branch.type === 'elseif') {
            if (branch.value === contextValue) {
              matchedContent = branch.content;
              matchedCondition = `${varName} = ${branch.value}`;
              break;
            }
          } else if (branch.type === 'else') {
            matchedContent = branch.content;
            matchedCondition = 'else (fallback)';
            break;
          }
        }
        
        if (!matchedContent && !matchedCondition) {
          return '';
        }
        
        if (matchedContent) {
          matchedContent = processConditionalsForPreview(matchedContent, context);
        }
        
        if (matchedCondition) {
          return `<div class="preview-condition-block"><div class="preview-condition-label">IF: ${matchedCondition}</div>${matchedContent}</div>`;
        }
        
        return matchedContent || '';
      } catch (err) {
        console.error('Error processing conditional:', err);
        return match;
      }
    });
  }
  
  // ============================================================
  // Open from Email Composer Context
  // ============================================================
  
  function openCurrentTemplateEditor() {
    console.log('openCurrentTemplateEditor called, templates:', airtableTemplates.length);
    
    const EMAIL_LINKS = window.EMAIL_LINKS || {};
    
    if (currentEmailContext) {
      const apptType = document.getElementById('emailApptType').value;
      const clientType = document.getElementById('emailClientType').value;
      const prepHandler = document.getElementById('emailPrepHandler').value;
      
      const brokerIntro = clientType === 'New' 
        ? `our Mortgage Broker ${currentEmailContext.broker}`
        : currentEmailContext.brokerFirst;
      
      templatePreviewContext = {
        greeting: currentEmailContext.greeting,
        broker: currentEmailContext.broker,
        brokerFirst: currentEmailContext.brokerFirst,
        brokerIntro: brokerIntro,
        appointmentTime: formatAppointmentTime(currentEmailContext.appointmentTime),
        appointmentType: apptType,
        clientType: clientType,
        prepHandler: prepHandler,
        daysUntil: calculateDaysUntil(currentEmailContext.appointmentTime),
        phoneNumber: currentEmailContext.phoneNumber,
        meetUrl: currentEmailContext.meetUrl,
        sender: currentEmailContext.sender,
        prefillNote: EMAIL_TEMPLATE.prefillNote[clientType] || '',
        officeMapLink: `<a href="${EMAIL_LINKS.officeMap || '#'}" target="_blank" style="color:#0066CC;">Office</a>`,
        ourTeamLink: `<a href="${EMAIL_LINKS.ourTeam || '#'}" target="_blank" style="color:#0066CC;">here</a>`,
        factFindLink: `<a href="${EMAIL_LINKS.factFind || '#'}" target="_blank" style="color:#0066CC;">Fact Find</a>`,
        myGovLink: `<a href="${EMAIL_LINKS.myGov || '#'}" target="_blank" style="color:#0066CC;">myGov</a>`,
        myGovVideoLink: `<a href="${EMAIL_LINKS.myGovVideo || '#'}" target="_blank" style="color:#0066CC;">myGov Video</a>`,
        incomeInstructionsLink: `<a href="${EMAIL_LINKS.incomeStatementInstructions || '#'}" target="_blank" style="color:#0066CC;">click here</a>`
      };
    } else {
      templatePreviewContext = null;
    }
    
    const confirmationTemplate = airtableTemplates.find(t => 
      t.type === 'Confirmation' || t.name.toLowerCase().includes('confirmation')
    );
    
    if (confirmationTemplate) {
      console.log('Opening confirmation template:', confirmationTemplate.id);
      openTemplateEditor(confirmationTemplate.id);
    } else if (airtableTemplates.length > 0) {
      console.log('Opening first template:', airtableTemplates[0].id);
      openTemplateEditor(airtableTemplates[0].id);
    } else {
      console.log('No templates, opening new');
      openTemplateEditor(null);
    }
  }
  
  // ============================================================
  // Utility Functions
  // ============================================================
  
  function formatAppointmentTime(dateStr) {
    if (!dateStr) return '[appointment time]';
    
    try {
      if (dateStr.includes('T') || dateStr.includes('Z')) {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;
        
        const options = {
          weekday: 'long',
          day: 'numeric',
          month: 'long',
          hour: 'numeric',
          minute: '2-digit',
          hour12: true,
          timeZone: 'Australia/Perth'
        };
        
        const formatted = date.toLocaleString('en-AU', options);
        
        const dayMatch = formatted.match(/(\d+)/);
        if (dayMatch) {
          const day = parseInt(dayMatch[1]);
          const suffix = getOrdinalSuffix(day);
          return formatted.replace(/(\d+)/, `${day}${suffix}`);
        }
        return formatted;
      }
      
      return dateStr;
    } catch (e) {
      console.error('Error formatting date:', e);
      return dateStr;
    }
  }
  
  function getOrdinalSuffix(day) {
    if (day > 3 && day < 21) return 'th';
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  }
  
  function calculateDaysUntil(appointmentTimeStr) {
    if (!appointmentTimeStr) return '?';
    
    try {
      let apptDate;
      
      if (appointmentTimeStr.includes('T') || appointmentTimeStr.includes('Z')) {
        apptDate = new Date(appointmentTimeStr);
        if (isNaN(apptDate.getTime())) return '?';
      } else {
        const months = {
          'january': 0, 'jan': 0, 'february': 1, 'feb': 1, 'march': 2, 'mar': 2,
          'april': 3, 'apr': 3, 'may': 4, 'june': 5, 'jun': 5, 'july': 6, 'jul': 6,
          'august': 7, 'aug': 7, 'september': 8, 'sep': 8, 'sept': 8,
          'october': 9, 'oct': 9, 'november': 10, 'nov': 10, 'december': 11, 'dec': 11
        };
        
        const cleanStr = appointmentTimeStr.toLowerCase().replace(/,/g, '');
        
        const dayMatch = cleanStr.match(/(\d{1,2})(st|nd|rd|th)?/);
        if (!dayMatch) return '?';
        const day = parseInt(dayMatch[1]);
        
        let monthIdx = -1;
        for (const [name, idx] of Object.entries(months)) {
          if (cleanStr.includes(name)) {
            monthIdx = idx;
            break;
          }
        }
        if (monthIdx === -1) return '?';
        
        const now = new Date();
        let year = now.getFullYear();
        apptDate = new Date(year, monthIdx, day);
        
        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        if (apptDate < today) {
          apptDate = new Date(year + 1, monthIdx, day);
        }
      }
      
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const apptDateOnly = new Date(apptDate.getFullYear(), apptDate.getMonth(), apptDate.getDate());
      
      const diffTime = apptDateOnly - today;
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      return diffDays >= 0 ? String(diffDays) : '?';
    } catch (e) {
      console.error('Error calculating days until:', e);
      return '?';
    }
  }
  
  // ============================================================
  // Gmail Integration
  // ============================================================
  
  function openInGmail() {
    if (!currentEmailContext) return;
    
    const to = document.getElementById('emailTo').value;
    const subject = document.getElementById('emailSubject').value;
    const previewEl = document.getElementById('emailPreviewBody');
    
    const body = convertHtmlToPlainTextWithUrls(previewEl.innerHTML);
    
    const gmailUrl = `https://mail.google.com/mail/?view=cm&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    
    window.open(gmailUrl, '_blank');
    closeEmailComposer();
  }
  
  function convertHtmlToPlainTextWithUrls(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    const links = temp.querySelectorAll('a');
    links.forEach(link => {
      const text = link.textContent;
      const href = link.getAttribute('href');
      const replacement = document.createTextNode(`${text} (${href})`);
      link.parentNode.replaceChild(replacement, link);
    });
    
    let result = temp.innerHTML;
    result = result.replace(/<br\s*\/?>/gi, '\n');
    
    const div = document.createElement('div');
    div.innerHTML = result;
    return div.textContent || div.innerText || '';
  }
  
  // ============================================================
  // Open Email Composer from Panel
  // ============================================================
  
  function openEmailComposerFromPanel(opportunityId) {
    const currentContactRecord = state.currentContactRecord;
    
    google.script.run.withSuccessHandler(function(oppData) {
      if (!oppData) {
        showAlert('Error', 'Could not load opportunity data', 'error');
        return;
      }
      const fields = oppData.fields || {};
      const contactFields = currentContactRecord ? currentContactRecord.fields : {};
      
      google.script.run.withSuccessHandler(function(appointments) {
        if (appointments && appointments.length > 0) {
          const scheduled = appointments.find(a => a.status === 'Scheduled' && a.appointmentTime);
          const appt = scheduled || appointments[0];
          if (appt) {
            fields._appointmentData = {
              appointmentTime: appt.appointmentTime,
              typeOfAppointment: appt.typeOfAppointment,
              phoneNumber: appt.phoneNumber,
              meetUrl: appt.meetUrl
            };
          }
        }
        
        const applicantIds = [];
        if (fields['Primary Applicant'] && fields['Primary Applicant'].length > 0) {
          applicantIds.push(...fields['Primary Applicant']);
        }
        if (fields['Applicants'] && fields['Applicants'].length > 0) {
          applicantIds.push(...fields['Applicants']);
        }
        
        if (applicantIds.length > 0) {
          let fetchedCount = 0;
          const emails = [];
          applicantIds.forEach(id => {
            google.script.run.withSuccessHandler(function(contact) {
              if (contact && contact.fields && contact.fields.EmailAddress1) {
                emails.push(contact.fields.EmailAddress1);
              }
              fetchedCount++;
              if (fetchedCount === applicantIds.length) {
                fields._applicantEmails = emails;
                openEmailComposer(fields, contactFields);
              }
            }).withFailureHandler(function() {
              fetchedCount++;
              if (fetchedCount === applicantIds.length) {
                fields._applicantEmails = emails;
                openEmailComposer(fields, contactFields);
              }
            }).getRecordById('Contacts', id);
          });
        } else {
          openEmailComposer(fields, contactFields);
        }
      }).withFailureHandler(function() {
        openEmailComposer(fields, contactFields);
      }).getAppointmentsForOpportunity(opportunityId);
    }).withFailureHandler(function(err) {
      showAlert('Error', 'Failed to load opportunity: ' + (err.message || 'Unknown error'), 'error');
    }).getRecordById('Opportunities', opportunityId);
  }
  
  // ============================================================
  // Initialize on Load
  // ============================================================
  
  loadEmailTemplates();
  
  // ============================================================
  // Expose Functions to Window
  // ============================================================
  
  window.openEmailComposer = openEmailComposer;
  window.closeEmailComposer = closeEmailComposer;
  window.sendEmail = sendEmail;
  window.updateEmailPreview = updateEmailPreview;
  window.openEmailComposerFromPanel = openEmailComposerFromPanel;
  window.openInGmail = openInGmail;
  
  window.loadEmailTemplates = loadEmailTemplates;
  window.openTemplateList = openTemplateList;
  window.closeTemplateList = closeTemplateList;
  window.refreshTemplateList = refreshTemplateList;
  window.createNewTemplate = createNewTemplate;
  window.seedDefaultTemplate = seedDefaultTemplate;
  
  window.openTemplateEditor = openTemplateEditor;
  window.closeTemplateEditor = closeTemplateEditor;
  window.saveTemplate = saveTemplate;
  window.openCurrentTemplateEditor = openCurrentTemplateEditor;
  
  window.insertVariable = insertVariable;
  window.insertConditionBlock = insertConditionBlock;
  window.updateConditionOptions = updateConditionOptions;
  window.updateTemplatePreviewFromControls = updateTemplatePreviewFromControls;
  
  window.formatAppointmentTime = formatAppointmentTime;
  window.calculateDaysUntil = calculateDaysUntil;

})();
