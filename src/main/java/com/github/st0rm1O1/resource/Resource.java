package com.github.st0rm1O1.resource;

import com.github.st0rm1O1.common.ImageRender;
import lombok.SneakyThrows;

import javax.imageio.ImageIO;
import javax.swing.plaf.PanelUI;
import java.awt.*;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.util.Objects;
import java.util.Random;


public class Resource {

	private Resource() {}

	private static final char separatorChar = '/';

	private static final String PATH_MIPMAP = separatorChar + "mipmap" + separatorChar;
	private static final String PATH_FONT = separatorChar + "font" + separatorChar;
	private static final String PATH_EXCEL = separatorChar + "excel" + separatorChar;

	private static final String EXCEL_PATH_TEST_DATA = PATH_EXCEL + "data.xlsx";

	private static final String FONT_PATH_INTER_REGULAR = PATH_FONT + "regular.ttf" ;
	private static final String FONT_PATH_INTER_MEDIUM = PATH_FONT + "medium.ttf" ;
	private static final String FONT_PATH_INTER_SEMIBOLD = PATH_FONT + "semibold.ttf" ;

	public static final String ICON_PATH_BOTTOM = PATH_MIPMAP + "bottom.png";
	public static final String ICON_PATH_CREATE = PATH_MIPMAP + "create.png";
	public static final String ICON_PATH_DELETE = PATH_MIPMAP + "delete.png";
	public static final String ICON_PATH_INSERT = PATH_MIPMAP + "insert.png";
	public static final String ICON_PATH_LEFT = PATH_MIPMAP + "left.png";
	public static final String ICON_PATH_RIGHT = PATH_MIPMAP + "right.png";
	public static final String ICON_PATH_TOP = PATH_MIPMAP + "top.png";
	public static final String ICON_PATH_UPDATE = PATH_MIPMAP + "update.png";
	public static final String ICON_PATH_ICON = PATH_MIPMAP + "icon.png";

	public static Font getInterRegular(int size) {
		return getInter(FONT_PATH_INTER_REGULAR, size);
	}

	public static Font getInterMedium(int size) {
		return getInter(FONT_PATH_INTER_MEDIUM, size);
	}

	public static Font getInterSemibold(int size) {
		return getInter(FONT_PATH_INTER_SEMIBOLD, size);
	}

	@SneakyThrows
	private static Font getInter(String path, int size) {
		return Font.createFont(Font.TRUETYPE_FONT, getResourceAsStream(path)).deriveFont(Font.PLAIN, size);
	}

	@SneakyThrows
	public static Image loadImage(String path) {
		return ImageIO.read(getResourceAsStream(path));
	}

	private static InputStream getResourceAsStream(String path) {
		InputStream input = Resource.class.getResourceAsStream(path);
		Objects.requireNonNull(input);
		return input;
	}

	@SneakyThrows
	private static URI getResourceAsURI(String path) {
		URL url = Resource.class.getResource(path);
		Objects.requireNonNull(url);
		return url.toURI();
	}

	public static PanelUI renderImage(String path) {
		return new ImageRender(path);
	}
}
