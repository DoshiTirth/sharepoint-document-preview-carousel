// SPFx's webpack config resolves imports of static asset files (images,
// fonts, and here .txt) to a string URL pointing at the deployed asset -
// same mechanism the project already relies on for icon/image assets.
declare module '*.txt' {
  const assetUrl: string;
  export default assetUrl;
}
