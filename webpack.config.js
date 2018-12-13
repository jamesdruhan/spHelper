const webpack      = require( 'webpack' );
const path         = require( 'path' );
// const MinifyPlugin = require( 'babel-minify-webpack-plugin' );

module.exports = {
	entry : 
	{
		"spHelper-stand-alone-poly.min.js" : ['babel-polyfill', './src/spHelper.js'],
		"spHelper-stand-alone.min.js"      : ['./src/spHelper.js'],
	},

	output :
	{
		path          : path.resolve(__dirname, 'build/'),
		filename      : '[name]',
		library       : "spHelper",
		libraryExport : "default",
	},

	// plugins :
	// [
    // 	new MinifyPlugin (),

	// 	new webpack.DefinePlugin
	// 	({
	// 		'process.env.NODE_ENV' : JSON.stringify( 'production' )
 	// 	}),
	// ],

	module :
	{
		rules :
		[
			// JS Processing
			{
				test    : /\.js$/,
				loader  : 'babel-loader',
				exclude : /node_modules/
			}
		]
	},

	stats :
	{
		colors : true
	},

	mode : 'development',
};
