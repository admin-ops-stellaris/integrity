document.addEventListener('DOMContentLoaded', async () => {
  const statusDiv = document.getElementById('status');
  
  try {
    const response = await fetch('/api/health');
    const data = await response.json();
    statusDiv.textContent = `API Status: ${data.app} — ${data.user ?? "unknown user"}`;
  } catch (error) {
    statusDiv.textContent = 'Unable to connect to API';
    statusDiv.classList.add('error');
  }
});
