import { navigateToSamePageHash } from '../SamePageHashNavigationHelper';

describe('navigateToSamePageHash', () => {
  afterEach(() => {
    document.body.innerHTML = '';
    if (window.location.hash) {
      window.history.replaceState(null, document.title, `${window.location.pathname}${window.location.search}`);
    }
  });

  it('scrolls the target when the requested hash is already current', () => {
    document.body.innerHTML = '<section id="section">Section</section>';
    window.history.replaceState(null, document.title, '#section');

    const targetSection = document.getElementById('section') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();

    const handled = navigateToSamePageHash(document, '#section', true);

    expect(handled).toBe(true);
    expect(targetSection.scrollIntoView).toHaveBeenCalledTimes(1);
  });
});
