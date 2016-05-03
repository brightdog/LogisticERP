<%
      ' 说明: ASP下的StringBuilder类
      ' 作者: Tank
      ' 日期: 2007-09-28
      ' 最后更新: 2009-02-28
      ' 描述: 用于快速组合大量字符

    Class ClsStringBuilder
        private sv_arrString()
        private sv_intArraryLength
        private sv_intCurrentLength
        private sv_intStep

        ' 初始化操作
        Private Sub Class_Initialize
            Step = 64
            Call Clear()
        End Sub
           
        ' 资源释放
        Private Sub class_terminate()
            Redim sv_arrString(0)
        End Sub

        ' 设置数组步进长度
        public property Let Step(Value)
            sv_intStep = Value
        end property

        ' 获取数组步进长度
        public property Get Step()
            Step = sv_intStep
        end property

        ' 返回数组长度
        public property Get Count()
            Count = sv_intCurrentLength
        end property

        ' 返回 splitter 分隔的字符串
        public function toString()
            if Count = 0 then
                toString = ""
                exit function
            end if

            ' 返回时重定义数组长度为实际长度
            if Count <= sv_intArraryLength then
                Redim Preserve sv_arrString(Count - 1)
                sv_intArraryLength = Count
            end if
            toString = join(sv_arrString, "")
        end function

        ' 插入字符串
        public function Append(byval s)
            sv_intCurrentLength = sv_intCurrentLength + 1
            if sv_intCurrentLength > sv_intArraryLength then
                sv_intArraryLength = sv_intArraryLength + Step
                Redim Preserve sv_arrString(sv_intArraryLength)
            end if

            sv_arrString(sv_intCurrentLength - 1) = s
        end function

        ' 清除所有字符串
        public function Clear()
            Redim sv_arrString(0)
            sv_intArraryLength = Step
            sv_intCurrentLength = 0
            Redim sv_arrString(sv_intArraryLength)
        end function
    End Class
%>