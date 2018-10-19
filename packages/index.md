# Packages

<ul class="packages">
  <li class="package">
    <h2>dictionary</h2>
    <p>
      Drop-in replacement for Scripting.Dictionary on Mac
    </p>
    <div>
      <a href="/packages/dictionary">Details</a>
    </div>
  </li>
  <li class="package">
    <h2>json</h2>
    <p>
      JSON conversion and parsing for VBA
    </p>
    <div>
      <a href="/packages/json">Details</a>
    </div>
  </li>
  <li class="package">
    <h2>web</h2>
    <p>
      Connect VBA, Excel, Access, and Office for Windows and Mac to web services and the web
    </p>
    <div>
      <a href="/packages/web">Details</a>
    </div>
  </li>
</ul>

<p class="message mod-highlight">
  More packages are coming soon and if you'd like to see your package here, <a href="mailto:tim@vba-blocks.com">please get in touch!</a>
</p>

<script type="text/javascript">
  const packages = document.querySelectorAll('.package');

  packages.forEach(function(package) {
    var link = package.querySelector('a');
    var down, up;

    // Guard for selecting text and other non-clicks
    // TODO handle ctrl+click
    package.addEventListener('mousedown', function(e) {
      down = +new Date();
    });
    package.addEventListener('mouseup', function(e) {
      up = +new Date();
      if ((up - down) < 200) {
        link.click();
      }
    })
  });
</script>
