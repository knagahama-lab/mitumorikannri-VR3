const fs = require('fs');

async function run() {
  const fetchContent = async (url) => {
    const r = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
    return await r.text();
  };
  try {
    console.log('Fetching CSS...');
    const css1 = await fetchContent('https://cdn.jsdelivr.net/npm/@tabler/core@1.0.0-beta19/dist/css/tabler.min.css');
    const css2 = await fetchContent('https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css');
    fs.writeFileSync('css.html', '<!-- Tabler CSS -->\n<style>\n' + css1 + '\n</style>\n<!-- DataTables CSS -->\n<style>\n' + css2 + '\n</style>\n');
    
    console.log('Fetching JS...');
    const js1 = await fetchContent('https://code.jquery.com/jquery-3.7.1.min.js');
    const js2 = await fetchContent('https://cdn.jsdelivr.net/npm/@tabler/core@1.0.0-beta19/dist/js/tabler.min.js');
    const js3 = await fetchContent('https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js');
    fs.writeFileSync('js.html', '<!-- jQuery -->\n<script>\n' + js1 + '\n</script>\n<!-- Tabler JS -->\n<script>\n' + js2 + '\n</script>\n<!-- DataTables JS -->\n<script>\n' + js3 + '\n</script>\n');
    console.log('Done.');
  } catch(e) {
    console.error(e);
  }
}
run();
