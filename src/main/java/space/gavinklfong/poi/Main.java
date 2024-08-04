package space.gavinklfong.poi;

import com.google.inject.Guice;
import com.google.inject.Injector;
import space.gavinklfong.poi.component.SimpleSheetExample;
import space.gavinklfong.poi.config.AppModule;

public class Main {

    public static void main(String[] args) {
        Injector injector = Guice.createInjector(new AppModule());
        SimpleSheetExample simpleSheetExample = injector.getInstance(SimpleSheetExample.class);
        simpleSheetExample.generateSimpleWorkbook("simple_sheet_example.xlsx");
        simpleSheetExample.generateWorkbookWithStyle("simple_sheet_style_example.xlsx");
    }
}