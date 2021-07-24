/**

        @RestController
        public class ExcelDown {

            @GetMapping("/excel/download")
            public void excelDownload(HttpServletResponse response) {
                List<Users> userList = getUsers(10);
                ExcelMaker excelMaker = new ExcelMaker();

                excelMaker.setSheetName("AAAAA")
                        .setRemoveField("id")
                        .setChangeFieldName("name", "이름")
                        .setChangeFieldName("phoneNumber", "핸드폰 번호")
                        .makeExcel(response, Users.class, userList, 1);
            }

            private List<User> getUsers(int index) {
                List<User> userList = new ArrayList<>();

                for (long i = 0; i < index; i++) {
                    Users users = new Users();
                    users.setId(i);
                    users.setName("LEE" + i);
                    users.setPhoneNumber("010-xxxx-xxx" + i);
                    userList.add(users);
                }

                return userList;
            }
        }


 */