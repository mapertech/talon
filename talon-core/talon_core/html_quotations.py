"""
The module's functions operate on message bodies trying to extract original
messages (without quoted messages) from html
"""

from __future__ import absolute_import
import regex as re

from lxml import etree

from talon_core.utils import cssselect 

CHECKPOINT_PREFIX = '#!%!'
CHECKPOINT_SUFFIX = '!%!#'
CHECKPOINT_PATTERN = re.compile(CHECKPOINT_PREFIX + r'\d+' + CHECKPOINT_SUFFIX)

# HTML quote indicators (tag ids)
QUOTE_IDS = [
    'OLK_SRC_BODY_SECTION',  # Outlook 2003
    'divRplyFwdMsg',         # Outlook 2010+ desktop/web reply marker
    'appendonsend',          # Outlook.com
    'divtagdefaultwrapper',  # Outlook.com alternate
]
RE_FWD = re.compile(r"^[-]+[ ]*Forwarded message[ ]*[-]+$", re.I | re.M)


def add_checkpoint(html_note, counter):
    """Recursively adds checkpoints to html tree.
    """
    if html_note.text:
        html_note.text = (html_note.text + CHECKPOINT_PREFIX +
                          str(counter) + CHECKPOINT_SUFFIX)
    else:
        html_note.text = (CHECKPOINT_PREFIX + str(counter) +
                          CHECKPOINT_SUFFIX)
    counter += 1

    for child in html_note.iterchildren():
        counter = add_checkpoint(child, counter)

    if html_note.tail:
        html_note.tail = (html_note.tail + CHECKPOINT_PREFIX +
                          str(counter) + CHECKPOINT_SUFFIX)
    else:
        html_note.tail = (CHECKPOINT_PREFIX + str(counter) +
                          CHECKPOINT_SUFFIX)
    counter += 1

    return counter


def delete_quotation_tags(html_note, counter, quotation_checkpoints):
    """Deletes tags with quotation checkpoints from html tree.
    """
    tag_in_quotation = True

    if quotation_checkpoints[counter]:
        html_note.text = ''
    else:
        tag_in_quotation = False
    counter += 1

    quotation_children = []  # Children tags which are in quotation.
    for child in html_note.iterchildren():
        counter, child_tag_in_quotation = delete_quotation_tags(
            child, counter,
            quotation_checkpoints
        )
        if child_tag_in_quotation:
            quotation_children.append(child)

    if quotation_checkpoints[counter]:
        html_note.tail = ''
    else:
        tag_in_quotation = False
    counter += 1

    if tag_in_quotation:
        return counter, tag_in_quotation
    else:
        # Remove quotation children.
        for child in quotation_children:
            html_note.remove(child)
        return counter, tag_in_quotation


def cut_gmail_quote(html_message):
    ''' Cuts the outermost block element with class gmail_quote. '''
    gmail_quote = cssselect('div.gmail_quote', html_message)
    if gmail_quote and (gmail_quote[0].text is None or not RE_FWD.match(gmail_quote[0].text)):
        gmail_quote[0].getparent().remove(gmail_quote[0])
        return True


def cut_microsoft_quote(html_message):
    ''' Cuts splitter block and all following blocks. '''
    #use EXSLT extensions to have a regex match() function with lxml
    ns = {"re": "http://exslt.org/regular-expressions"}

    #general pattern: @style='border:none;border-top:solid <color> 1.0pt;padding:3.0pt 0<unit> 0<unit> 0<unit>'
    #outlook 2007, 2010 (international) <color=#B5C4DF> <unit=cm>
    #outlook 2007, 2010 (american)      <color=#B5C4DF> <unit=pt>
    #outlook 2013       (international) <color=#E1E1E1> <unit=cm>
    #outlook 2013       (american)      <color=#E1E1E1> <unit=pt>
    #also handles a variant with a space after the semicolon
    splitter = html_message.xpath(
        #outlook 2007, 2010, 2013 (international, american)
        "//div[@style[re:match(., 'border:none; ?border-top:solid #(E1E1E1|B5C4DF) 1.0pt; ?"
        "padding:3.0pt 0(in|cm) 0(in|cm) 0(in|cm)')]]|"
        #windows mail
        "//div[@style='padding-top: 5px; "
        "border-top-color: rgb(229, 229, 229); "
        "border-top-width: 1px; border-top-style: solid;']"
        , namespaces=ns
    )

    if splitter:
        splitter = splitter[0]
        #outlook 2010
        if splitter == splitter.getparent().getchildren()[0]:
            splitter = splitter.getparent()
    else:
        #outlook 2003
        splitter = html_message.xpath(
            "//div"
            "/div[@class='MsoNormal' and @align='center' "
            "and @style='text-align:center']"
            "/font"
            "/span"
            "/hr[@size='3' and @width='100%' and @align='center' "
            "and @tabindex='-1']"
        )
        if len(splitter):
            splitter = splitter[0]
            splitter = splitter.getparent().getparent()
            splitter = splitter.getparent().getparent()

    if len(splitter):
        parent = splitter.getparent()
        after_splitter = splitter.getnext()
        while after_splitter is not None:
            parent.remove(after_splitter)
            after_splitter = splitter.getnext()
        parent.remove(splitter)
        return True

    return False


def cut_by_id(html_message):
    """Remove elements matching known reply-marker ids and (for Outlook reply
    markers) every sibling that follows the marker — that's where the quoted
    body lives in Outlook's standard structure."""
    found = False
    for quote_id in QUOTE_IDS:
        quote = cssselect('#{}'.format(quote_id), html_message)
        if not quote:
            continue
        found = True
        marker = quote[0]
        parent = marker.getparent()
        if parent is None:
            continue

        # Outlook puts the visual <hr> separator right before divRplyFwdMsg —
        # drop it too if present.
        prev = marker.getprevious()
        if prev is not None and prev.tag == 'hr':
            parent.remove(prev)

        # Remove the marker and everything after it in the parent.
        for sib in list(marker.itersiblings()):
            parent.remove(sib)
        parent.remove(marker)
    return found


def cut_blockquote(html_message):
    ''' Cuts the last non-nested blockquote with wrapping elements.'''
    quote = html_message.xpath(
        '(.//blockquote)'
        '[not(@class="gmail_quote") and not(ancestor::blockquote)]'
        '[last()]')

    if quote:
        quote = quote[0]
        quote.getparent().remove(quote)
        return True


_FROM_PREFIXES = ('From:', 'Date:', 'De:', 'Fecha:', 'Data:')


def _outermost_from_matches(html_message, mode):
    """Return outermost elements whose text_content (mode='text') or .tail
    (mode='tail') starts with one of the From: prefixes.

    Uses a Python tree walk rather than lxml XPath because Outlook's
    Word-exported HTML carries unbound xmlns:o/v/w/m prefixes that crash
    lxml's XPath engine. The walk is namespace-tolerant.

    'Outermost' means we filter out matches that are descendants of other
    matches — so for <p><b><span>De:</span></b>...</p> we return the <p>,
    not the <span>.
    """
    raw = []
    for el in html_message.iter():
        # Skip Comment / ProcessingInstruction nodes (their .tag is a callable,
        # not a string) — itertext() raises on them.
        if not isinstance(el.tag, str):
            continue
        if mode == 'text':
            # Concatenate all text descendants. lxml's _Element has no
            # .text_content(); itertext() is the etree-native equivalent.
            text = ''.join(el.itertext())
        else:
            text = el.tail or ''
        if text and text.strip().startswith(_FROM_PREFIXES):
            raw.append(el)
    raw_ids = {id(e) for e in raw}
    outermost = []
    for el in raw:
        anc = el.getparent()
        while anc is not None and id(anc) not in raw_ids:
            anc = anc.getparent()
        if anc is None:
            outermost.append(el)
    return outermost


def cut_from_block(html_message):
    """Cuts div tag which wraps block starting with "From:"."""
    # Case 1: From: block is enclosed in some tag — find the outermost matching
    # element, walk up to a div ancestor, and remove the div + everything after.
    block = _outermost_from_matches(html_message, mode='text')

    if block:
        # Use the FIRST match (earliest in document order) as the cut point so
        # the removal cascades through every header paragraph and the quoted
        # body below. Using the last match leaves earlier headers behind.
        target = block[0]
        parent_div = None
        node = target
        while node.getparent() is not None:
            if node.tag == 'div':
                parent_div = node
                break
            node = node.getparent()
        if parent_div is None:
            return False

        maybe_body = parent_div.getparent()
        parent_div_is_all_content = (
            maybe_body is not None and maybe_body.tag == 'body' and
            len(maybe_body.getchildren()) == 1)

        if not parent_div_is_all_content:
            # Remove parent_div + every sibling that follows it.
            parent = parent_div.getparent()
            for sib in list(parent_div.itersiblings()):
                parent.remove(sib)
            parent.remove(parent_div)
            return True

        # Outlook Word-export wraps content in a single <div class="WordSection1">
        # whose direct children are <p class="MsoNormal"><b>De:</b>…</p> paragraphs.
        # Don't remove the whole div — strip the matched paragraph and every
        # paragraph after it within the same parent.
        parent = target.getparent()
        if parent is not None:
            for sib in list(target.itersiblings()):
                parent.remove(sib)
            parent.remove(target)
            return True
        return False

    # Case 2: From: block goes right after e.g. <hr> as the next-sibling's tail.
    block = _outermost_from_matches(html_message, mode='tail')
    if block:
        first = block[0]
        if RE_FWD.match(first.getparent().text or ''):
            return False
        while first.getnext() is not None:
            first.getparent().remove(first.getnext())
        first.getparent().remove(first)
        return True

def cut_zimbra_quote(html_message):
    zDivider = html_message.xpath('//hr[@data-marker="__DIVIDER__"]')
    if zDivider:
        zDivider[0].getparent().remove(zDivider[0])
        return True
