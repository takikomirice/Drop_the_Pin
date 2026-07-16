module.exports = Object.freeze({
  allowed: Object.freeze([
    'https://example.com',
    'http://example.com/path',
    'http://intranet/path',
    'https://example.com:8443/path',
    'https://127.0.0.1/path',
    'https://[2001:db8::1]/path',
    'https://[::ffff:192.0.2.1]/path',
    'https://example.com/a%20b'
  ]),
  rejected: Object.freeze([
    'javascript:alert(1)',
    'ftp://example.com',
    'https://',
    'https://example.com:99999',
    'http://999.999.999.999',
    'https://exa^mple.com',
    'https://exa|mple.com',
    'https://%zz.com',
    'https://%00.com',
    'https://example.com/\u0001control'
  ])
});
