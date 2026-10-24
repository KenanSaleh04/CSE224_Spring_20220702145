CC = gcc
LEX = flex
YACC = bison -y -d
B = build
IN = ../inputs
OUT = ../outputs

all: t1 t2 t3 calc3 t7 bonus

t1:
	mkdir -p $(B)/t1
	for f in ex1 ex2 ex3 ex3mod ex4; do \
	    $(LEX) -o $(B)/t1/$$f.c t1/$$f.l && \
	    $(CC) $(B)/t1/$$f.c -o $(B)/t1/$$f -lfl || exit 1; \
	done

t2:
	mkdir -p $(B)/t2
	cd $(B)/t2 && $(YACC) ../../t2/calc1.y && $(LEX) ../../t2/calc1.l && \
	    $(CC) y.tab.c lex.yy.c -o calc1

t3:
	mkdir -p $(B)/t3
	cd $(B)/t3 && $(YACC) ../../t3/calc2.y && $(LEX) ../../t3/calc2.l && \
	    $(CC) y.tab.c lex.yy.c -o calc2

calc3:
	mkdir -p $(B)/calc3
	cd $(B)/calc3 && $(YACC) ../../calc3/calc3.y && $(LEX) ../../calc3/calc3.l && \
	    $(CC) -I../../calc3 -c y.tab.c lex.yy.c && \
	    $(CC) -I. -I../../calc3 y.tab.o lex.yy.o ../../calc3/calc3a.c -o calc3a && \
	    $(CC) -I. -I../../calc3 y.tab.o lex.yy.o ../../calc3/calc3b.c -o calc3b && \
	    $(CC) -I. -I../../calc3 y.tab.o lex.yy.o ../../calc3/calc3g.c -o calc3g

t7:
	mkdir -p $(B)/t7np $(B)/t7t
	cd $(B)/t7np && bison -y -d -v ../../t7/calc3np.y && $(LEX) ../../calc3/calc3.l && \
	    $(CC) -I. -I../../calc3 y.tab.c lex.yy.c ../../calc3/calc3a.c -o calc3np
	cd $(B)/t7t && bison -y -d -t ../../t7/calc3t.y && $(LEX) ../../calc3/calc3.l && \
	    $(CC) -I. -I../../calc3 y.tab.c lex.yy.c ../../calc3/calc3a.c -o calc3t

bonus:
	mkdir -p $(B)/bonus
	cd $(B)/bonus && $(YACC) ../../bonus/calc4.y && $(LEX) ../../bonus/calc4.l && \
	    $(CC) -I../../bonus -c y.tab.c lex.yy.c && \
	    $(CC) -I. -I../../bonus y.tab.o lex.yy.o ../../bonus/calc4a.c -o calc4a && \
	    $(CC) -I. -I../../bonus y.tab.o lex.yy.o ../../bonus/calc4b.c -o calc4b && \
	    $(CC) -I. -I../../bonus y.tab.o lex.yy.o ../../bonus/calc4g.c -o calc4g

define runit
	@echo "$$ ./$(notdir $(1)) < inputs/$(2).txt" > $(OUT)/$(3).txt
	@stdbuf -oL $(B)/$(1) < $(IN)/$(2).txt >> $(OUT)/$(3).txt 2>&1 || true
endef

run: all
	mkdir -p $(OUT)
	$(call runit,t1/ex1,t1_text,t1_ex1_text)
	$(call runit,t1/ex1,t1_nomatch,t1_ex1_nomatch)
	@echo "$$ ./ex2 inputs/t1_text.txt" > $(OUT)/t1_ex2_text.txt
	@$(B)/t1/ex2 $(IN)/t1_text.txt >> $(OUT)/t1_ex2_text.txt
	@echo "$$ ./ex2 inputs/t1_lastline.txt" > $(OUT)/t1_ex2_lastline.txt
	@$(B)/t1/ex2 $(IN)/t1_lastline.txt >> $(OUT)/t1_ex2_lastline.txt
	$(call runit,t1/ex3,t1_text,t1_ex3_text)
	$(call runit,t1/ex3,t1_nomatch,t1_ex3_nomatch)
	$(call runit,t1/ex3,t1_code,t1_ex3_code)
	$(call runit,t1/ex3mod,t1_code,t1_ex3mod_code)
	$(call runit,t1/ex4,t1_text,t1_ex4_text)
	$(call runit,t1/ex4,t1_nomatch,t1_ex4_nomatch)
	$(call runit,t1/ex4,t1_lastline,t1_ex4_lastline)
	$(call runit,t2/calc1,t2_given1,t2_given1)
	$(call runit,t2/calc1,t2_given2,t2_given2)
	$(call runit,t2/calc1,t2_given3,t2_given3)
	$(call runit,t2/calc1,t2_own1,t2_own1)
	$(call runit,t2/calc1,t2_own2,t2_own2)
	$(call runit,t2/calc1,t2_own3,t2_own3)
	$(call runit,t3/calc2,t3_given,t3_given)
	$(call runit,t3/calc2,t3_own1,t3_own1)
	$(call runit,t3/calc2,t3_own2,t3_own2)
	$(call runit,t3/calc2,t3_own3,t3_own3)
	$(call runit,calc3/calc3a,t4_test,t4_test)
	$(call runit,calc3/calc3a,t4_own1,t4_own1)
	$(call runit,calc3/calc3a,t4_own2,t4_own2)
	$(call runit,calc3/calc3a,t4_own3,t4_own3)
	$(call runit,calc3/calc3b,t4_test,t5_test)
	$(call runit,calc3/calc3b,t4_own1,t5_own1)
	$(call runit,calc3/calc3b,t4_own2,t5_own2)
	$(call runit,calc3/calc3b,t4_own3,t5_own3)
	$(call runit,calc3/calc3g,t6_given,t6_given)
	$(call runit,calc3/calc3g,t6_own1,t6_own1)
	$(call runit,calc3/calc3g,t6_own2,t6_own2)
	$(call runit,calc3/calc3g,t6_own3,t6_own3)
	$(call runit,t7np/calc3np,t7_given,t7_noprec_given)
	$(call runit,t7np/calc3np,t7_own,t7_noprec_own)
	$(call runit,calc3/calc3a,t7_own,t7_prec_own)
	$(call runit,t7t/calc3t,t7_given,t7_trace)
	cp $(B)/t7np/y.output $(OUT)/t7_y.output
	$(call runit,bonus/calc4a,bonus_test,bonus_interp)
	$(call runit,bonus/calc4b,bonus_test,bonus_compile)
	$(call runit,bonus/calc4a,bonus_err,bonus_err)
	$(call runit,bonus/calc4g,bonus_graph,bonus_graph)

clean:
	rm -rf $(B)

.PHONY: all t1 t2 t3 calc3 t7 bonus run clean
