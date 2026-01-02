document.addEventListener('DOMContentLoaded', () => {
  const articleContent = document.querySelector('.article-content');
  if (!articleContent) return;

  const headers = articleContent.querySelectorAll('h2, h3');
  if (headers.length === 0) return;

  // Create TOC Container
  const tocContainer = document.createElement('div');
  tocContainer.className = 'toc-container';

  // Header
  const tocHeader = document.createElement('div');
  tocHeader.className = 'toc-header';
  
  const tocTitle = document.createElement('span');
  tocTitle.className = 'toc-title';
  tocTitle.textContent = '快速導覽';
  
  const tocToggle = document.createElement('button');
  tocToggle.className = 'toc-toggle';
  tocToggle.ariaLabel = '切換導覽';
  tocToggle.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <line x1="8" y1="6" x2="21" y2="6"></line>
      <line x1="8" y1="12" x2="21" y2="12"></line>
      <line x1="8" y1="18" x2="21" y2="18"></line>
      <line x1="3" y1="6" x2="3.01" y2="6"></line>
      <line x1="3" y1="12" x2="3.01" y2="12"></line>
      <line x1="3" y1="18" x2="3.01" y2="18"></line>
    </svg>
  `;

  tocHeader.appendChild(tocTitle);
  tocHeader.appendChild(tocToggle);
  tocContainer.appendChild(tocHeader);

  // Content
  const tocContent = document.createElement('div');
  tocContent.className = 'toc-content';
  const tocList = document.createElement('ul');
  
  let currentH2Item = null;
  let h2Sublist = null;

  headers.forEach((header, index) => {
    // Generate ID
    if (!header.id) {
      header.id = 'toc-heading-' + index;
    }

    const listItem = document.createElement('li');
    const link = document.createElement('a');
    link.href = '#' + header.id;
    link.textContent = header.textContent;
    
    // Smooth scroll
    link.addEventListener('click', (e) => {
      e.preventDefault();
      header.scrollIntoView({ behavior: 'smooth' });
    });

    if (header.tagName.toLowerCase() === 'h2') {
      listItem.appendChild(link);
      tocList.appendChild(listItem);
      currentH2Item = listItem;
      h2Sublist = null; // Reset sublist for new H2
    } else if (header.tagName.toLowerCase() === 'h3') {
      // If we have an H2 parent, append to its sublist
      if (currentH2Item) {
        if (!h2Sublist) {
          h2Sublist = document.createElement('ul');
          currentH2Item.appendChild(h2Sublist);
        }
        listItem.appendChild(link);
        h2Sublist.appendChild(listItem);
      } else {
        // Orphan H3 (shouldn't happen usually, but handle it)
        listItem.appendChild(link);
        tocList.appendChild(listItem);
      }
    }
  });

  tocContent.appendChild(tocList);
  tocContainer.appendChild(tocContent);

  // Insert TOC before the first header or at top
  const firstHeader = headers[0];
  if (firstHeader) {
    articleContent.insertBefore(tocContainer, firstHeader);
  } else {
    articleContent.prepend(tocContainer);
  }

  // Toggle functionality
  let isOpen = true;
  tocToggle.addEventListener('click', () => {
    isOpen = !isOpen;
    tocContent.style.display = isOpen ? 'block' : 'none';
    tocContainer.classList.toggle('is-collapsed', !isOpen);
  });
});
