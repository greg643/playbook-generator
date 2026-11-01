// Syntax guard only. GSS still verifies the signature, expiry and account;
// Team reads separately verify current membership. Never decode for authority.
// GSS mint_token uses urlsafe_b64encode and retains trailing '=' padding.
export const validGSSBearer = value => typeof value === 'string' &&
  value.length >= 16 && value.length <= 16384 && !/\s/.test(value) && /^[A-Za-z0-9._~-]+={0,2}$/.test(value);
