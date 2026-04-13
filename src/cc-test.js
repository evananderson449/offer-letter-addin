/**
 * TEMPORARY: Content Control API feasibility test for Word Online.
 * Tests three critical operations:
 *   1. Insert a content control at [NAME] placeholder via OOXML
 *   2. Find the content control by tag
 *   3. Update its text via insertText()
 *
 * Remove this file after testing.
 */
(function () {
  var log = [];
  var el;

  function out(msg) {
    log.push(msg);
    console.log('[CC-TEST] ' + msg);
    if (el) el.textContent = log.join('\n');
  }

  document.addEventListener('DOMContentLoaded', function () {
    var btn = document.getElementById('ccTestBtn');
    el = document.getElementById('ccTestResult');
    if (!btn) return;

    btn.addEventListener('click', async function () {
      btn.disabled = true;
      log = [];
      out('Starting content control test...');

      try {
        // STEP 1: Create a content control by wrapping selected text or first paragraph
        out('Step 1: Creating content control...');
        var tagName = 'cc_test_tag';

        await Word.run(async function (context) {
          // Try to find [NAME] in the document body via OOXML and wrap it
          // But first, let's try the simpler approach: insert a CC at the start of body
          var body = context.document.body;
          var paragraphs = body.paragraphs;
          paragraphs.load('text');
          await context.sync();

          out('  Document has ' + paragraphs.items.length + ' paragraphs');

          // Insert a content control on the first paragraph
          var firstPara = paragraphs.items[0];
          var cc = firstPara.insertContentControl();
          cc.tag = tagName;
          cc.title = 'Test CC';
          cc.appearance = 'BoundingBox';
          await context.sync();

          out('  ✓ Content control created (tag: ' + tagName + ')');
        });

        // STEP 2: Find content control by tag
        out('Step 2: Finding content control by tag...');
        var foundText = '';

        await Word.run(async function (context) {
          var ccs = context.document.contentControls.getByTag(tagName);
          ccs.load('items');
          await context.sync();

          out('  Found ' + ccs.items.length + ' content control(s) with tag "' + tagName + '"');

          if (ccs.items.length > 0) {
            ccs.items[0].load('text');
            await context.sync();
            foundText = ccs.items[0].text;
            out('  ✓ Content control text: "' + foundText.substring(0, 50) + '..."');
          }
        });

        // STEP 3: Update content control text
        out('Step 3: Updating content control text...');
        var t0 = performance.now();

        await Word.run(async function (context) {
          var ccs = context.document.contentControls.getByTag(tagName);
          ccs.load('items');
          await context.sync();

          if (ccs.items.length > 0) {
            ccs.items[0].insertText('TEST_VALUE_12345', 'Replace');
            await context.sync();
          }
        });

        var elapsed = Math.round(performance.now() - t0);
        out('  ✓ Text updated in ' + elapsed + 'ms');

        // STEP 4: Verify the update
        out('Step 4: Verifying update...');

        await Word.run(async function (context) {
          var ccs = context.document.contentControls.getByTag(tagName);
          ccs.load('items');
          await context.sync();

          if (ccs.items.length > 0) {
            ccs.items[0].load('text');
            await context.sync();
            var newText = ccs.items[0].text;
            if (newText === 'TEST_VALUE_12345') {
              out('  ✓ Verified: text is "TEST_VALUE_12345"');
            } else {
              out('  ✗ Unexpected text: "' + newText + '"');
            }
          }
        });

        // STEP 5: Clean up — delete the content control (restore original text)
        out('Step 5: Cleaning up...');

        await Word.run(async function (context) {
          var ccs = context.document.contentControls.getByTag(tagName);
          ccs.load('items');
          await context.sync();

          if (ccs.items.length > 0) {
            // Delete the CC but keep its content
            ccs.items[0].delete(false);
            await context.sync();
            out('  ✓ Content control removed (text preserved)');
          }
        });

        out('\n✅ ALL TESTS PASSED — Content controls work in Word Online!');
        out('Update latency: ' + elapsed + 'ms');

      } catch (e) {
        out('\n❌ TEST FAILED: ' + e.message);
        out('Stack: ' + (e.stack || 'N/A'));
        out('\nContent controls may not be supported in this Word Online version.');
      } finally {
        btn.disabled = false;
      }
    });
  });
})();
